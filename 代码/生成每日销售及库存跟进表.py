# -*- coding: utf-8 -*-
"""
生成【数据中间表/每日数据中间表/每日数据中间表_D.xlsx】并更新【结果表/每日销售及库存跟进表.xlsx】。

VSCode 直接运行：
    1. 将下方 VSCODE_RUN_CONFIG["enabled"] 改为 True。
    2. 修改 VSCODE_RUN_CONFIG["root"] 和 VSCODE_RUN_CONFIG["date"]。
    3. 在 VSCode 中打开本文件，点击 Run Python File。

重要说明：
1. 执行日期 D 由 --date 指定；商品销售统计使用 D-1 文件，其余源数据使用 D 文件。
   累计销售源数据按 D 的日期位置选择月初至 D-1 的累计口径。
2. 每日数据中间表会被保留，且核心指标写入 Excel 公式。
3. 结果表先按数据粒度表校准行集合和排序，再按粒度键同步中间表数据。
4. 中间表生成后会由脚本调用 Excel 打开、全量计算、保存并关闭；结果表读取中间表公式计算后的缓存值。
"""

import argparse
import gc
import os
import re
import subprocess
import sys
import tempfile
import time
from copy import copy
from datetime import datetime, timedelta
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter


VSCODE_RUN_CONFIG = {
    # True：忽略命令行参数，直接使用这里的 root/date，适合 VSCode 右上角 Run Python File。
    # False：使用命令行 --root 和 --date 参数。
    "enabled": True,
    "root": r"<ROOT>",
    "date": "2026-06-30",
}


SHEET_SKU_MID = "SKU日销-中间"
SHEET_CAT_MID = "类目日销-中间"
SHEET_OP_MID = "运营日销-中间"
SHEET_C1_MID = "C1运营类目目标跟踪-中间表"
SHEET_C2_MID = "C2运营店铺目标跟踪-中间表"
SHEET_SKU_SCOPE = "SKU日销数据口径"
SHEET_CAT_SCOPE = "类目日销数据口径"
SHEET_OP_SCOPE = "运营日销数据口径"
SHEET_C1_SCOPE = "C1运营类目目标跟踪数据口径"
SHEET_C2_SCOPE = "C2运营店铺目标跟踪数据口径"
SHEET_SALES = "千易商品销售统计"
SHEET_SALES_CUM = "千易商品销售统计_累计"
SHEET_STOCK = "千易库存数据"
SHEET_IN_NEW = "千易入库单_新建"
SHEET_IN_RECEIVING = "千易入库单_收货中"
SHEET_FEISHU = "飞书-入库单跟进表"
SHEET_SKU_MAP = "SKU信息表"
SHEET_STORE_MAP = "店铺信息表"
SHEET_WAREHOUSE_MAP = "仓库信息表"

RESULT_SKU = "SKU日销"
RESULT_CAT = "类目日销"
RESULT_OP = "运营日销"
RESULT_C1 = "C1运营类目目标跟踪"
RESULT_C2 = "C2运营店铺目标跟踪"
GRANULARITY_SKU = "SKU日销（管理）"
GRANULARITY_CAT = "类目日销（管理）"
GRANULARITY_OP = "运营日销"
GRANULARITY_C1 = "C1运营类目目标跟踪"
GRANULARITY_C2 = "C2运营店铺目标跟踪"

HEADER_FILL = PatternFill("solid", fgColor="E8EEF7")
FORMULA_COLUMN_FILL = PatternFill("solid", fgColor="DDEBF7")


def verify_file_released(path) :
    if not path.exists():
        return
    temp_path = path.with_name(f".{path.name}.lockcheck.{os.getpid()}.tmp")
    if temp_path.exists():
        temp_path.unlink()
    try:
        path.rename(temp_path)
        temp_path.rename(path)
    except PermissionError as exc:
        raise PermissionError(f"文件仍被 Python 或 Excel 占用，请关闭后重试：{path}") from exc
    finally:
        if temp_path.exists() and not path.exists():
            temp_path.rename(path)


def release_resources_and_verify(paths) :
    gc.collect()
    time.sleep(0.5)
    for path in paths:
        verify_file_released(path)
    log_stage("已释放脚本内工作簿资源，并完成输出文件占用校验")


SALES_HEADERS = [
    "系统sku", "sku名称", "spu 编码", "spu", "spu值", "主SKU分类", "币种", "平台",
    "店铺", "运营", "站点", "站点(CN)", "仓库", "日均销量", "销售额", "销售订单量",
    "销售sku数量", "销售sku数量VALUE", "销售SKU数量(不限)", "发货前取消订单量",
    "发货前取消金额", "取消sku数量", "发货后退款订单量", "退款sku数量", "平均售价",
    "平均sku数量", "手工单量", "手工单sku数量", "补发货单量", "补发货sku数量",
    "时间类型", "时间范围",
]

STOCK_HEADERS = [
    "商品编码", "商品名称", "主SKU分类", "仓库", "站点（中文）", "站点（英文）",
    "可用量", "在途量", "未收到库存", "实际在途量", "库存合计", "3天销量",
    "7天销量", "15天销量", "30天销量", "60天销量", "90天销量",
    "新建入库单数量", "总在途（含新建入库单）",
]

IN_NEW_HEADERS = [
    "入库单号", "收货仓", "站点（中文）", "站点（英文）", "商品编码", "预计数量",
    "预计数量VALUE", "发货时间", "到仓时间", "到仓时间整合", "到货天数",
]

FEISHU_HEADERS = ["仓库", "站点", "入库单号", "发货时间", "到仓时间", "到仓时间整合", "是否到货", "到货天数"]

SKU_SCOPE_HEADERS = ["站点(中文)", "站点(英文)", "产品分级", "产品主类/材质", "主SKU分类", "产品编码", "产品名称"]
CAT_SCOPE_HEADERS = ["站点(中文)", "站点(英文)", "产品分级", "产品主类/材质", "主SKU分类"]
OP_SCOPE_HEADERS = ["站点(中文)", "站点(英文)", "产品分级", "产品主类/材质", "主SKU分类", "运营"]
C1_SCOPE_HEADERS = ["站点", "产品分级", "产品主类/材质", "主SKU分类", "运营"]
C2_SCOPE_HEADERS = ["站点", "店铺名", "运营", "产品分级", "产品主类/材质", "主SKU分类"]


def sku_mid_headers(sales_label) :
    return [
        "供应商", "站点", "产品分级", "产品主类/材质", "主SKU分类", "SKU", "货值单价",
        "货值+海运单价", "销售额单价", "账单收入单价", "当日库存", "月销量",
        "可售月数", "30天内到货在途数量", "30天货值", "30天货值+海运",
        "30天销售额", "30天账单收入", "60天内到货在途数量", "60天货值",
        "60天货值+海运", "60天销售额", "60天账单收入", "90天在途数量",
        "90天货值", "90天货值+海运", "90天销售额", "90天账单收入",
        "库存状况", sales_label,
    ]


def cat_mid_headers(sales_label) :
    return [
        "供应商", "站点", "产品分级", "产品主类/材质", "主SKU分类", "当日库存", "月销量",
        "可售月数", "30天内到货在途数量", "30天货值", "30天货值+海运",
        "30天销售额", "30天账单收入", "60天内到货在途数量", "60天货值",
        "60天货值+海运", "60天销售额", "60天账单收入", "90天在途数量",
        "90天货值", "90天货值+海运", "90天销售额", "90天账单收入",
        "库存状况", sales_label,
    ]


def op_mid_headers(sales_label) :
    return ["站点", "产品分级", "产品主类/材质", "主SKU分类", "运营", sales_label]


C1_RESULT_BASE_HEADERS = C1_SCOPE_HEADERS
C2_RESULT_BASE_HEADERS = C2_SCOPE_HEADERS


def c1_mid_headers(cumulative_label) :
    return C1_SCOPE_HEADERS + [cumulative_label]


def c2_mid_headers(cumulative_label) :
    return C2_SCOPE_HEADERS + [cumulative_label]


SKU_RESULT_BASE_HEADERS = [
    "供应商", "站点", "产品分级", "产品主类/材质", "主SKU分类", "SKU", "产品名称", "货值单价",
    "货值+海运单价", "销售额单价", "账单收入单价", "当日库存", "月销量",
    "可售月数", "30天内到货在途数量", "30天货值", "30天货值+海运",
    "30天销售额", "30天账单收入", "60天内到货在途数量", "60天货值",
    "60天货值+海运", "60天销售额", "60天账单收入", "90天在途数量",
    "90天货值", "90天货值+海运", "90天销售额", "90天账单收入", "库存状况",
]

CAT_RESULT_BASE_HEADERS = [
    "供应商", "站点", "产品分级", "产品主类/材质", "主SKU分类", "当日库存", "月销量",
    "可售月数", "30天内到货在途数量", "30天货值", "30天货值+海运",
    "30天销售额", "30天账单收入", "60天内到货在途数量", "60天货值",
    "60天货值+海运", "60天销售额", "60天账单收入", "90天在途数量",
    "90天货值", "90天货值+海运", "90天销售额", "90天账单收入", "库存状况",
]

OP_RESULT_BASE_HEADERS = ["站点", "产品分级", "产品主类/材质", "主SKU分类", "运营"]
SKU_RESULT_DIMENSION_HEADERS = {"站点", "产品分级", "产品主类/材质", "主SKU分类", "SKU", "产品名称"}
CAT_RESULT_DIMENSION_HEADERS = {"站点", "产品分级", "产品主类/材质", "主SKU分类"}
OP_RESULT_DIMENSION_HEADERS = {"站点", "产品分级", "产品主类/材质", "主SKU分类", "运营"}
SKU_RESULT_PRICE_SYNC_HEADERS = {"货值单价", "货值+海运单价", "销售额单价", "账单收入单价"}
SKU_RESULT_DERIVED_FORMULA_HEADERS = {
    "30天货值", "30天货值+海运", "30天销售额", "30天账单收入",
    "60天货值", "60天货值+海运", "60天销售额", "60天账单收入",
    "90天货值", "90天货值+海运", "90天销售额", "90天账单收入",
}
SKU_RESULT_FORMULA_HEADERS = SKU_RESULT_DERIVED_FORMULA_HEADERS
SKU_RESULT_DATA_HEADERS = [
    header
    for header in SKU_RESULT_BASE_HEADERS
    if header not in SKU_RESULT_DIMENSION_HEADERS and header not in SKU_RESULT_FORMULA_HEADERS
]
CAT_RESULT_DATA_HEADERS = [header for header in CAT_RESULT_BASE_HEADERS if header not in CAT_RESULT_DIMENSION_HEADERS]
OP_RESULT_DATA_HEADERS = [header for header in OP_RESULT_BASE_HEADERS if header not in OP_RESULT_DIMENSION_HEADERS]
SKU_RESULT_CLEAR_HEADERS = [header for header in SKU_RESULT_DATA_HEADERS if header != "库存状况"]
CAT_RESULT_SUM_FORMULA_HEADERS = [
    header for header in CAT_RESULT_DATA_HEADERS if header not in {"供应商", "可售月数", "库存状况"}
]
CAT_RESULT_FORMULA_HEADERS = set(CAT_RESULT_SUM_FORMULA_HEADERS) | {"可售月数"}
RESULT_BLANK_FIELDS = {"产品分级", "产品主类/材质", "库存状况"}


class RunContext:
    """保存一次每日任务所需的运行上下文。"""

    def __init__(
        self,
        root,
        run_date,
        sales_date,
        sales_label,
        growth_label,
        cumulative_label,
        middle_path,
        result_path,
    ):
        self.root = root
        self.run_date = run_date
        self.sales_date = sales_date
        self.sales_label = sales_label
        self.growth_label = growth_label
        self.cumulative_label = cumulative_label
        self.middle_path = middle_path
        self.result_path = result_path


def parse_args() :
    parser = argparse.ArgumentParser(description="生成每日数据中间表并更新每日销售及库存跟进结果表")
    parser.add_argument("--root", required=True, help="根目录，目录下必须包含 源数据/数据中间表/结果表")
    parser.add_argument("--date", required=True, help="执行日期 D，格式 YYYY-MM-DD")
    return parser.parse_args()


def resolve_run_parameters(root = None, date = None) :
    """解析运行参数。

    优先级：
    1. 函数显式传入 root/date，适合其他 Python 文件 import 后调用。
    2. VSCODE_RUN_CONFIG["enabled"] 为 True 时，使用配置块。
    3. 否则读取命令行 --root/--date。
    """
    if root and date:
        return root, date
    if VSCODE_RUN_CONFIG.get("enabled"):
        config_root = str(VSCODE_RUN_CONFIG.get("root") or "").strip()
        config_date = str(VSCODE_RUN_CONFIG.get("date") or "").strip()
        if not config_root or not config_date:
            raise ValueError("VSCODE_RUN_CONFIG 已启用，但 root 或 date 为空")
        print("使用 VSCODE_RUN_CONFIG 参数运行")
        print(f"root = {config_root}")
        print(f"date = {config_date}")
        return config_root, config_date
    args = parse_args()
    return args.root, args.date


def build_context(root_text, date_text) :
    root = Path(root_text).expanduser().resolve()
    if not root.exists():
        raise FileNotFoundError(f"根目录不存在：{root}")
    if not (root / "源数据").exists():
        raise FileNotFoundError(f"源数据目录不存在：{root / '源数据'}")
    (root / "数据中间表" / "每日数据中间表").mkdir(parents=True, exist_ok=True)
    (root / "结果表").mkdir(parents=True, exist_ok=True)

    try:
        run_date = datetime.strptime(date_text, "%Y-%m-%d")
    except ValueError as exc:
        raise ValueError("执行日期必须为 YYYY-MM-DD 格式") from exc

    sales_date = run_date - timedelta(days=1)
    sales_label = sales_date.strftime("%Y.%m.%d") + "销量"
    growth_label = sales_date.strftime("%Y.%m.%d") + "涨幅"
    cumulative_label = sales_date.strftime("%Y.%m.%d") + "累计销量"
    middle_path = root / "数据中间表" / "每日数据中间表" / f"每日数据中间表_{run_date:%Y-%m-%d}.xlsx"
    result_path = root / "结果表" / "每日销售及库存跟进表.xlsx"
    return RunContext(root, run_date, sales_date, sales_label, growth_label, cumulative_label, middle_path, result_path)


def cumulative_sales_path(ctx) :
    sales_dir = ctx.root / "源数据" / "千易销售数据"
    if ctx.run_date.day == 2:
        return sales_dir / f"商品销售统计_{ctx.sales_date:%Y-%m-%d}.xlsx"
    if ctx.run_date.day == 1:
        return sales_dir / f"商品销售统计_{ctx.sales_date:%Y-%m}.xlsx"
    month_start = ctx.run_date.replace(day=1)
    return sales_dir / f"商品销售统计_累计_{month_start:%Y-%m-%d}_{ctx.sales_date:%Y-%m-%d}.xlsx"


def source_paths(ctx) :
    d = ctx.run_date.strftime("%Y-%m-%d")
    s = ctx.sales_date.strftime("%Y-%m-%d")
    return {
        "sales": ctx.root / "源数据" / "千易销售数据" / f"商品销售统计_{s}.xlsx",
        "sales_cum": cumulative_sales_path(ctx),
        "stock": ctx.root / "源数据" / "千易库存数据" / f"库存查询_{d}.xlsx",
        "in_new": ctx.root / "源数据" / "千易入库单数据" / f"入库单明细_新建_{d}.xlsx",
        "in_receiving": ctx.root / "源数据" / "千易入库单数据" / f"入库单明细_收货中_{d}.xlsx",
        "feishu": ctx.root / "源数据" / "飞书人工维护数据" / "入库单跟踪表" / f"入库单跟踪表_{d}.xlsx",
        "mapping": ctx.root / "源数据" / "映射表" / "SKU、店铺、仓库映射表.xlsx",
        "granularity": find_granularity_path(ctx.root),
    }


def find_granularity_path(root) :
    candidates = [
        path
        for path in (root / "源数据").rglob("@数据粒度表_每日销售及库存跟进表*.xlsx")
        if not path.name.startswith("~$")
    ]
    if not candidates:
        raise FileNotFoundError("未找到源数据中的 @数据粒度表_每日销售及库存跟进表*.xlsx")
    return max(candidates, key=lambda path: (path.stat().st_mtime, path.name))


def validate_sources(paths) :
    missing = [str(path) for path in paths.values() if not path.exists()]
    if missing:
        raise FileNotFoundError("以下源文件不存在：\n" + "\n".join(missing))
    empty = []
    for name, path in paths.items():
        if name == "granularity":
            continue
        # 部分千易导出的 xlsx 会把 worksheet dimension 元数据错误写成 A1；
        # read_only 模式会信任该元数据并误判为空，因此源文件有效性校验使用普通模式读取。
        wb = load_workbook(path, read_only=False, data_only=False)
        try:
            ws = wb[wb.sheetnames[0]]
            has_data = False
            for row in ws.iter_rows(min_row=2, max_row=min(ws.max_row, 20), values_only=True):
                if any(value not in (None, "") for value in row):
                    has_data = True
                    break
            if ws.max_row <= 1 or not has_data:
                empty.append(f"{name}: {path}")
        finally:
            wb.close()
    if empty:
        raise ValueError("以下源文件疑似为空或只有单个表头，已停止执行：\n" + "\n".join(empty))
    validate_granularity_source(paths["granularity"])


def validate_granularity_source(path) :
    wb = load_workbook(path, read_only=True, data_only=True)
    requirements = {
        GRANULARITY_SKU: ["站点(英文)", "产品编码", "产品分级", "产品主类/材质", "主SKU分类", "产品名称"],
        GRANULARITY_CAT: ["站点(英文)", "主SKU分类", "产品分级", "产品主类/材质"],
        GRANULARITY_OP: ["站点(英文)", "主SKU分类", "运营", "产品分级", "产品主类/材质"],
        GRANULARITY_C1: C1_SCOPE_HEADERS,
        GRANULARITY_C2: C2_SCOPE_HEADERS,
    }
    missing_sheets = set(requirements) - set(wb.sheetnames)
    if missing_sheets:
        raise KeyError(f"数据粒度表缺少 sheet：{sorted(missing_sheets)}")
    for sheet_name, headers in requirements.items():
        ws = wb[sheet_name]
        for header in headers:
            find_header_col(ws, header, row=2)
        if ws.max_row < 3:
            raise ValueError(f"数据粒度表 {sheet_name} 没有有效粒度行")
    wb.close()


def first_sheet(path) :
    wb = load_workbook(path, data_only=False)
    return wb[wb.sheetnames[0]]


def remove_default_sheet(wb) :
    if wb.sheetnames == ["Sheet"]:
        del wb["Sheet"]


def ensure_sheet(wb, name) :
    if name in wb.sheetnames:
        ws = wb[name]
        ws.delete_rows(1, ws.max_row)
        return ws
    return wb.create_sheet(name)


def normalize_header(value) :
    return str(value).replace("\xa0", "").strip() if value is not None else ""


def header_map(ws, row = 1) :
    mapping = {}
    for col in range(1, ws.max_column + 1):
        value = ws.cell(row, col).value
        if value is not None and str(value) not in mapping:
            mapping[str(value)] = col
    return mapping


def find_header_col(ws, header, row = 1) :
    target = normalize_header(header)
    for col in range(1, ws.max_column + 1):
        if normalize_header(ws.cell(row, col).value) == target:
            return col
    raise KeyError(f"sheet {ws.title} 缺少字段：{header}")


def find_header_col_containing(ws, header_text, row = 1) :
    target = re.sub(r"\s+", "", normalize_header(header_text))
    for col in range(1, ws.max_column + 1):
        candidate = re.sub(r"\s+", "", normalize_header(ws.cell(row, col).value))
        if target and target in candidate:
            return col
    raise KeyError(f"sheet {ws.title} 缺少包含文本的字段：{header_text}")


def ensure_header_col(ws, header) :
    try:
        return find_header_col(ws, header)
    except KeyError:
        col = ws.max_column + 1
        cell = ws.cell(1, col, header)
        cell.font = Font(bold=True)
        return col


def write_headers(ws, headers) :
    for col, header in enumerate(headers, 1):
        cell = ws.cell(1, col, header)
        cell.font = Font(bold=True)
        cell.fill = copy(HEADER_FILL)
    ws.freeze_panes = "A2"


def apply_formula_column_fill(ws, columns) :
    for col in columns:
        for row in range(1, ws.max_row + 1):
            ws.cell(row, col).fill = copy(FORMULA_COLUMN_FILL)


def copy_cell_style(src, dst) :
    if src.has_style:
        dst.font = copy(src.font)
        dst.fill = copy(src.fill)
        dst.border = copy(src.border)
        dst.alignment = copy(src.alignment)
        dst.number_format = src.number_format
        dst.protection = copy(src.protection)


def copy_sheet_values(src, dst) :
    for row in src.iter_rows():
        for cell in row:
            target = dst.cell(cell.row, cell.column, cell.value)
            copy_cell_style(cell, target)
    for key, dim in src.column_dimensions.items():
        dst.column_dimensions[key].width = dim.width


def copy_source_to_headers(
    src,
    dst,
    headers,
    header_row = 1,
    data_start_row = None,
) :
    write_headers(dst, headers)
    source_header = header_map(src, header_row)
    first_data_row = data_start_row or header_row + 1
    for target_col, header in enumerate(headers, 1):
        source_col = source_header.get(header)
        if not source_col:
            continue
        for source_row in range(first_data_row, src.max_row + 1):
            target_row = source_row - first_data_row + 2
            dst.cell(target_row, target_col, src.cell(source_row, source_col).value)


def copy_mapping_sheet(mapping_path, source_sheet, target_wb, target_sheet) :
    source_wb = load_workbook(mapping_path, data_only=False)
    try:
        if source_sheet not in source_wb.sheetnames:
            raise KeyError(f"映射表缺少 sheet：{source_sheet}")
        ws = ensure_sheet(target_wb, target_sheet)
        copy_sheet_values(source_wb[source_sheet], ws)
    finally:
        source_wb.close()


def col_letter(col) :
    return get_column_letter(col)


def as_number_formula(cell_ref) :
    return f"={cell_ref}+0"


def log_stage(message) :
    print(f"[阶段完成] {message}")


def log_sheet_done(ws, label = None) :
    name = label or ws.title
    data_rows = max(ws.max_row - 1, 0)
    print(f"[Sheet完成] {name}：数据行={data_rows}，字段数={ws.max_column}")


def recalculate_workbook_with_excel(path, timeout_seconds = 600, workbook_label = "中间表"):
    """Use desktop Excel to calculate formulas and save cached results for openpyxl data_only reads."""
    workbook_path = str(path.resolve())
    log_stage(f"开始调用 Excel 计算{workbook_label}公式：{path}")
    try:
        import win32com.client  # type: ignore

        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        workbook = None
        try:
            workbook = excel.Workbooks.Open(workbook_path, UpdateLinks=0, ReadOnly=False)
            excel.CalculateFullRebuild()
            started = time.time()
            while excel.CalculationState != 0:
                if time.time() - started > timeout_seconds:
                    raise TimeoutError("Excel 公式计算超时")
                time.sleep(0.5)
            workbook.Save()
        finally:
            if workbook is not None:
                workbook.Close(SaveChanges=True)
            excel.Quit()
        log_stage(f"已完成 Excel 公式计算并保存{workbook_label}")
        return
    except ImportError:
        pass

    powershell_script = r"""
param([string]$WorkbookPath, [int]$TimeoutSeconds)
$ErrorActionPreference = 'Stop'
$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $workbook = $excel.Workbooks.Open($WorkbookPath, 0, $false)
    $excel.CalculateFullRebuild()
    $start = Get-Date
    while ($excel.CalculationState -ne 0) {
        if (((Get-Date) - $start).TotalSeconds -gt $TimeoutSeconds) {
            throw "Excel 公式计算超时"
        }
        Start-Sleep -Milliseconds 500
    }
    $workbook.Save()
}
finally {
    if ($workbook -ne $null) {
        $workbook.Close($true) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }
    if ($excel -ne $null) {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
"""
    script_path = None
    try:
        with tempfile.NamedTemporaryFile("w", suffix=".ps1", delete=False, encoding="utf-8-sig") as script_file:
            script_file.write(powershell_script)
            script_path = Path(script_file.name)
        completed = subprocess.run(
            [
                "powershell",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(script_path),
                "-WorkbookPath",
                workbook_path,
                "-TimeoutSeconds",
                str(timeout_seconds),
            ],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout_seconds + 60,
        )
        if completed.returncode != 0:
            raise RuntimeError(
                f"调用 Excel 计算{workbook_label}公式失败，请确认本机已安装 Microsoft Excel 且文件未被占用。\n"
                + completed.stderr.strip()
            )
    finally:
        if script_path:
            try:
                script_path.unlink()
            except FileNotFoundError:
                pass
    log_stage(f"已完成 Excel 公式计算并保存{workbook_label}")


def excel_external_prefix(path, sheet_name) :
    directory = str(path.parent.resolve()).replace("/", "\\")
    return f"'{directory}\\[{path.name}]{sheet_name}'!"


def external_cell(path, sheet_name, cell_ref) :
    return f"={excel_external_prefix(path, sheet_name)}{cell_ref}"


def external_range(path, sheet_name, range_ref) :
    return f"{excel_external_prefix(path, sheet_name)}{range_ref}"


def row_has_any_value(ws, row, cols = None) :
    cols = cols or range(1, ws.max_column + 1)
    return any(ws.cell(row, col).value not in (None, "") for col in cols)


def unique_rows_by_key(ws, headers, key_headers) :
    hmap = header_map(ws)
    keys_seen = set()
    output = []
    for row in range(2, ws.max_row + 1):
        item = {header: ws.cell(row, hmap[header]).value if header in hmap else None for header in headers}
        key = tuple(item.get(k) for k in key_headers)
        if any(v in (None, "") for v in key):
            continue
        if key in keys_seen:
            continue
        keys_seen.add(key)
        output.append(item)
    return output


def build_lookup(ws, key_header, value_header) :
    hmap = header_map(ws)
    if key_header not in hmap or value_header not in hmap:
        raise KeyError(f"sheet {ws.title} 缺少字段：{key_header} 或 {value_header}")
    lookup = {}
    for row in range(2, ws.max_row + 1):
        key = ws.cell(row, hmap[key_header]).value
        value = ws.cell(row, hmap[value_header]).value
        if key not in (None, "") and key not in lookup:
            lookup[key] = value
    return lookup


def sales_scope_rows(
    sales_ws,
    headers,
    key_headers,
    sku_main_map,
    store_operator_map,
    store_site_cn_map,
) :
    hmap = header_map(sales_ws)
    required = ["系统sku", "sku名称", "店铺", "站点", "仓库"]
    missing = [header for header in required if header not in hmap]
    if missing:
        raise KeyError(f"{SHEET_SALES} 缺少字段：" + "、".join(missing))

    rows = []
    seen = set()
    for row in range(2, sales_ws.max_row + 1):
        sku = sales_ws.cell(row, hmap["系统sku"]).value
        store = sales_ws.cell(row, hmap["店铺"]).value
        item = {
            "系统sku": sku,
            "sku名称": sales_ws.cell(row, hmap["sku名称"]).value,
            "主SKU分类": sku_main_map.get(sku),
            "店铺": store,
            "运营": store_operator_map.get(store),
            "站点": sales_ws.cell(row, hmap["站点"]).value,
            "站点(CN)": store_site_cn_map.get(store),
            "仓库": sales_ws.cell(row, hmap["仓库"]).value,
        }
        key = tuple(item.get(header) for header in key_headers)
        if any(value in (None, "") for value in key):
            continue
        if key in seen:
            continue
        seen.add(key)
        rows.append({header: item.get(header) for header in headers})
    return rows


def create_sales_sheet(wb, source_ws, sheet_name = SHEET_SALES) :
    ws = ensure_sheet(wb, sheet_name)
    copy_source_to_headers(source_ws, ws, SALES_HEADERS)
    for row in range(2, ws.max_row + 1):
        ws.cell(row, 6, f"=VLOOKUP(A{row},{SHEET_SKU_MAP}!F:L,7,0)")
        ws.cell(row, 10, f"=VLOOKUP(I{row},{SHEET_STORE_MAP}!B:D,2,0)")
        ws.cell(row, 12, f"=VLOOKUP(I{row},{SHEET_STORE_MAP}!B:D,3,0)")
        ws.cell(row, 18, as_number_formula(f"Q{row}"))
    apply_formula_column_fill(ws, [6, 10, 12, 18])
    return ws


def create_in_receiving_sheet(wb, source_ws) :
    ws = ensure_sheet(wb, SHEET_IN_RECEIVING)
    copy_sheet_values(source_ws, ws)
    write_headers_if_missing(ws)
    expected_col = find_header_col(ws, "预计数量")
    shelved_col = find_header_col(ws, "仓库上架数量")
    remark_col = find_header_col(ws, "批注")
    unreceived_col = ensure_header_col(ws, "未收到库存")
    expected_letter = col_letter(expected_col)
    shelved_letter = col_letter(shelved_col)
    remark_letter = col_letter(remark_col)
    for row in range(2, ws.max_row + 1):
        if row_has_any_value(ws, row):
            ws.cell(
                row,
                unreceived_col,
                f'=IF(ISNUMBER(FIND("点完",{remark_letter}{row})),MAX({expected_letter}{row}-{shelved_letter}{row},0),0)',
            )
    apply_formula_column_fill(ws, [unreceived_col])
    return ws


def write_headers_if_missing(ws) :
    for col in range(1, ws.max_column + 1):
        if ws.cell(1, col).value is None:
            ws.cell(1, col, "")


def create_feishu_sheet(wb, source_ws, sales_date) :
    ws = ensure_sheet(wb, SHEET_FEISHU)
    copy_source_to_headers(source_ws, ws, FEISHU_HEADERS)
    for row in range(2, ws.max_row + 1):
        if row_has_any_value(ws, row):
            ws.cell(row, 8, f'=IFERROR(F{row}-DATE({sales_date.year},{sales_date.month},{sales_date.day}),"")')
    apply_formula_column_fill(ws, [8])
    return ws


def create_in_new_sheet(wb, source_ws) :
    ws = ensure_sheet(wb, SHEET_IN_NEW)
    copy_source_to_headers(source_ws, ws, IN_NEW_HEADERS)
    for row in range(2, ws.max_row + 1):
        if not row_has_any_value(ws, row, [1, 2, 5, 6]):
            continue
        ws.cell(row, 3, f"=VLOOKUP(B{row},{SHEET_WAREHOUSE_MAP}!A:C,2,0)")
        ws.cell(row, 4, f"=VLOOKUP(B{row},{SHEET_WAREHOUSE_MAP}!A:C,3,0)")
        ws.cell(row, 7, as_number_formula(f"F{row}"))
        ws.cell(row, 8, f"=VLOOKUP(A{row},'{SHEET_FEISHU}'!C:H,2,0)")
        ws.cell(row, 9, f"=VLOOKUP(A{row},'{SHEET_FEISHU}'!C:H,3,0)")
        ws.cell(row, 10, f"=VLOOKUP(A{row},'{SHEET_FEISHU}'!C:H,4,0)")
        ws.cell(row, 11, f"=IFERROR(VLOOKUP(A{row},'{SHEET_FEISHU}'!C:H,6,0),-1)")
    apply_formula_column_fill(ws, [3, 4, 7, 8, 9, 10, 11])
    return ws


def create_stock_sheet(
    wb,
    source_ws,
    receiving_ws,
    in_new_ws,
) :
    ws = ensure_sheet(wb, SHEET_STOCK)
    copy_source_to_headers(source_ws, ws, STOCK_HEADERS, header_row=2, data_start_row=3)
    receiving_sku_col = col_letter(find_header_col(receiving_ws, "商品编码"))
    receiving_warehouse_col = col_letter(find_header_col(receiving_ws, "收货仓"))
    receiving_unreceived_col = col_letter(find_header_col(receiving_ws, "未收到库存"))
    in_new_sku_col = col_letter(find_header_col(in_new_ws, "商品编码"))
    in_new_warehouse_col = col_letter(find_header_col(in_new_ws, "收货仓"))
    in_new_quantity_col = col_letter(find_header_col(in_new_ws, "预计数量VALUE"))
    stock_sku_letter = col_letter(find_header_col(ws, "商品编码"))
    stock_warehouse_letter = col_letter(find_header_col(ws, "仓库"))
    new_quantity_col = find_header_col(ws, "新建入库单数量")
    total_transit_col = find_header_col(ws, "总在途（含新建入库单）")
    actual_transit_col = find_header_col(ws, "实际在途量")
    new_quantity_letter = col_letter(new_quantity_col)
    actual_transit_letter = col_letter(actual_transit_col)
    for row in range(2, ws.max_row + 1):
        if not row_has_any_value(ws, row, [1, 2, 4]):
            continue
        ws.cell(row, 3, f"=VLOOKUP(A{row},{SHEET_SKU_MAP}!F:L,7,0)")
        ws.cell(row, 5, f"=VLOOKUP(D{row},{SHEET_WAREHOUSE_MAP}!A:C,2,0)")
        ws.cell(row, 6, f"=VLOOKUP(D{row},{SHEET_WAREHOUSE_MAP}!A:C,3,0)")
        ws.cell(
            row,
            9,
            f"=SUMIFS('{SHEET_IN_RECEIVING}'!{receiving_unreceived_col}:{receiving_unreceived_col},"
            f"'{SHEET_IN_RECEIVING}'!{receiving_sku_col}:{receiving_sku_col},{stock_sku_letter}{row},"
            f"'{SHEET_IN_RECEIVING}'!{receiving_warehouse_col}:{receiving_warehouse_col},{stock_warehouse_letter}{row})",
        )
        ws.cell(row, 10, f"=IF((H{row}-I{row})<0,0,H{row}-I{row})")
        ws.cell(row, 11, f"=J{row}+G{row}")
        ws.cell(
            row,
            new_quantity_col,
            f"=SUMIFS('{SHEET_IN_NEW}'!{in_new_quantity_col}:{in_new_quantity_col},"
            f"'{SHEET_IN_NEW}'!{in_new_warehouse_col}:{in_new_warehouse_col},{stock_warehouse_letter}{row},"
            f"'{SHEET_IN_NEW}'!{in_new_sku_col}:{in_new_sku_col},{stock_sku_letter}{row})",
        )
        ws.cell(row, total_transit_col, f"={actual_transit_letter}{row}+{new_quantity_letter}{row}")
    apply_formula_column_fill(ws, [3, 5, 6, 9, 10, 11, new_quantity_col, total_transit_col])
    return ws


def create_scope_sheet(
    wb,
    source_ws,
    sheet_name,
    headers,
    key_headers,
    sku_main_map,
    store_operator_map,
    store_site_cn_map,
) :
    ws = ensure_sheet(wb, sheet_name)
    write_headers(ws, headers)
    rows = sales_scope_rows(source_ws, headers, key_headers, sku_main_map, store_operator_map, store_site_cn_map)
    for out_row, item in enumerate(rows, 2):
        for col, header in enumerate(headers, 1):
            ws.cell(out_row, col, item.get(header))
    return ws


def create_sku_mid_sheet(wb, sales_label) :
    ws = ensure_sheet(wb, SHEET_SKU_MID)
    headers = sku_mid_headers(sales_label)
    write_headers(ws, headers)
    scope = wb[SHEET_SKU_SCOPE]
    scope_hmap = header_map(scope)
    for row in range(2, scope.max_row + 1):
        out_row = row
        station = scope.cell(row, scope_hmap["站点(英文)"]).value
        main_sku = scope.cell(row, scope_hmap["主SKU分类"]).value
        sku = scope.cell(row, scope_hmap["产品编码"]).value
        if not station or not sku:
            continue
        ws.cell(out_row, 1, f"=_xlfn.XLOOKUP(F{out_row},{SHEET_SKU_MAP}!F:F,{SHEET_SKU_MAP}!A:A)")
        ws.cell(out_row, 2, station)
        ws.cell(out_row, 3, None)  # 产品分级：置空
        ws.cell(out_row, 4, f"=_xlfn.XLOOKUP(F{out_row},{SHEET_SKU_MAP}!F:F,{SHEET_SKU_MAP}!K:K)")
        ws.cell(out_row, 5, main_sku)
        ws.cell(out_row, 6, sku)
        ws.cell(out_row, 7, f"=_xlfn.XLOOKUP(F{out_row},{SHEET_SKU_MAP}!$F:$F,{SHEET_SKU_MAP}!$M:$M)")
        ws.cell(out_row, 8, f'=IFERROR(_xlfn.IFS(B{out_row}="PH",_xlfn.XLOOKUP(F{out_row},{SHEET_SKU_MAP}!$F:$F,{SHEET_SKU_MAP}!$O:$O),B{out_row}="ID",_xlfn.XLOOKUP(F{out_row},{SHEET_SKU_MAP}!$F:$F,{SHEET_SKU_MAP}!$R:$R),B{out_row}="TH",_xlfn.XLOOKUP(F{out_row},{SHEET_SKU_MAP}!$F:$F,{SHEET_SKU_MAP}!$W:$W),B{out_row}="MY",_xlfn.XLOOKUP(F{out_row},{SHEET_SKU_MAP}!$F:$F,{SHEET_SKU_MAP}!$AB:$AB),TRUE,"未知站点"),0)')
        ws.cell(out_row, 9, f"=J{out_row}*1.3")
        ws.cell(out_row, 10, f'=IFERROR(_xlfn.IFS(OR(B{out_row}="PH",B{out_row}="MY",B{out_row}="TH"),H{out_row}*1.25,B{out_row}="ID",H{out_row}*1.45,TRUE,0),0)')
        ws.cell(out_row, 11, f"=SUMIFS({SHEET_STOCK}!G:G,{SHEET_STOCK}!A:A,F{out_row},{SHEET_STOCK}!F:F,B{out_row})")
        ws.cell(out_row, 12, f"=SUMIFS({SHEET_STOCK}!O:O,{SHEET_STOCK}!A:A,F{out_row},{SHEET_STOCK}!F:F,B{out_row})")
        ws.cell(out_row, 13, f"=IFERROR(K{out_row}/L{out_row},0)")
        ws.cell(out_row, 14, f"=SUMIFS({SHEET_STOCK}!J:J,{SHEET_STOCK}!A:A,F{out_row},{SHEET_STOCK}!F:F,B{out_row})")
        ws.cell(out_row, 15, f"=N{out_row}*G{out_row}")
        ws.cell(out_row, 16, f"=N{out_row}*H{out_row}")
        ws.cell(out_row, 17, f"=N{out_row}*I{out_row}")
        ws.cell(out_row, 18, f"=N{out_row}*J{out_row}")
        ws.cell(out_row, 19, f'=SUMIFS({SHEET_IN_NEW}!G:G,{SHEET_IN_NEW}!K:K,">=0",{SHEET_IN_NEW}!K:K,"<=60",{SHEET_IN_NEW}!D:D,B{out_row},{SHEET_IN_NEW}!E:E,F{out_row})')
        ws.cell(out_row, 20, f"=S{out_row}*G{out_row}")
        ws.cell(out_row, 21, f"=S{out_row}*H{out_row}")
        ws.cell(out_row, 22, f"=S{out_row}*I{out_row}")
        ws.cell(out_row, 23, f"=S{out_row}*J{out_row}")
        ws.cell(out_row, 24, f'=SUMIFS({SHEET_IN_NEW}!G:G,{SHEET_IN_NEW}!K:K,">60",{SHEET_IN_NEW}!D:D,B{out_row},{SHEET_IN_NEW}!E:E,F{out_row})')
        ws.cell(out_row, 25, f"=X{out_row}*G{out_row}")
        ws.cell(out_row, 26, f"=X{out_row}*H{out_row}")
        ws.cell(out_row, 27, f"=X{out_row}*I{out_row}")
        ws.cell(out_row, 28, f"=X{out_row}*J{out_row}")
        ws.cell(out_row, 29, None)  # 库存状况：置空
        ws.cell(out_row, 30, f"=SUMIFS({SHEET_SALES}!$R:$R,{SHEET_SALES}!$K:$K,$B{out_row},{SHEET_SALES}!$A:$A,$F{out_row})")
    return ws


def create_cat_mid_sheet(wb, sales_label) :
    ws = ensure_sheet(wb, SHEET_CAT_MID)
    headers = cat_mid_headers(sales_label)
    write_headers(ws, headers)
    scope = wb[SHEET_CAT_SCOPE]
    scope_hmap = header_map(scope)
    for row in range(2, scope.max_row + 1):
        out_row = row
        station = scope.cell(row, scope_hmap["站点(英文)"]).value
        main_sku = scope.cell(row, scope_hmap["主SKU分类"]).value
        if not station or not main_sku:
            continue
        ws.cell(out_row, 1, f"=_xlfn.XLOOKUP(E{out_row},{SHEET_SKU_MAP}!L:L,{SHEET_SKU_MAP}!A:A)")
        ws.cell(out_row, 2, station)
        ws.cell(out_row, 3, None)
        ws.cell(out_row, 4, f"=_xlfn.XLOOKUP(E{out_row},{SHEET_SKU_MAP}!L:L,{SHEET_SKU_MAP}!K:K)")
        ws.cell(out_row, 5, main_sku)
        sumifs_cols = {
            6: "K", 7: "L", 9: "N", 10: "O", 11: "P", 12: "Q", 13: "R",
            14: "S", 15: "T", 16: "U", 17: "V", 18: "W",
            19: "X", 20: "Y", 21: "Z", 22: "AA", 23: "AB", 25: "AD",
        }
        for target_col, sku_col in sumifs_cols.items():
            ws.cell(out_row, target_col, f"=SUMIFS('{SHEET_SKU_MID}'!{sku_col}:{sku_col},'{SHEET_SKU_MID}'!$B:$B,$B{out_row},'{SHEET_SKU_MID}'!$E:$E,$E{out_row})")
        ws.cell(out_row, 8, f"=IFERROR(F{out_row}/G{out_row},0)")
        ws.cell(out_row, 24, None)
    return ws


def create_op_mid_sheet(wb, sales_label) :
    ws = ensure_sheet(wb, SHEET_OP_MID)
    headers = op_mid_headers(sales_label)
    write_headers(ws, headers)
    scope = wb[SHEET_OP_SCOPE]
    scope_hmap = header_map(scope)
    for row in range(2, scope.max_row + 1):
        out_row = row
        station = scope.cell(row, scope_hmap["站点(英文)"]).value
        main_sku = scope.cell(row, scope_hmap["主SKU分类"]).value
        operator = scope.cell(row, scope_hmap["运营"]).value
        if not station or not main_sku or not operator:
            continue
        ws.cell(out_row, 1, station)
        ws.cell(out_row, 2, None)
        ws.cell(out_row, 3, f"=_xlfn.XLOOKUP(D{out_row},{SHEET_SKU_MAP}!L:L,{SHEET_SKU_MAP}!K:K)")
        ws.cell(out_row, 4, main_sku)
        ws.cell(out_row, 5, operator)
        ws.cell(out_row, 6, f"=SUMIFS({SHEET_SALES}!$R:$R,{SHEET_SALES}!$K:$K,$A{out_row},{SHEET_SALES}!$J:$J,$E{out_row},{SHEET_SALES}!$F:$F,$D{out_row})")
    return ws


def create_c2_mid_sheet(wb, cumulative_label) :
    ws = ensure_sheet(wb, SHEET_C2_MID)
    write_headers(ws, c2_mid_headers(cumulative_label))
    scope = wb[SHEET_C2_SCOPE]
    scope_hmap = header_map(scope)
    for row in range(2, scope.max_row + 1):
        out_row = row
        station = scope.cell(row, scope_hmap["站点"]).value
        store = scope.cell(row, scope_hmap["店铺名"]).value
        operator = scope.cell(row, scope_hmap["运营"]).value
        sku_category = scope.cell(row, scope_hmap["主SKU分类"]).value
        if not station or not store or not operator or not sku_category:
            continue
        for col, header in enumerate(C2_SCOPE_HEADERS, 1):
            ws.cell(out_row, col, scope.cell(row, scope_hmap[header]).value)
        ws.cell(
            out_row,
            7,
            f"=SUMIFS('{SHEET_SALES_CUM}'!$R:$R,"
            f"'{SHEET_SALES_CUM}'!$L:$L,$A{out_row},"
            f"'{SHEET_SALES_CUM}'!$I:$I,$B{out_row},"
            f"'{SHEET_SALES_CUM}'!$F:$F,$F{out_row})",
        )
    apply_formula_column_fill(ws, [7])
    return ws


def create_c1_mid_sheet(wb, cumulative_label) :
    ws = ensure_sheet(wb, SHEET_C1_MID)
    write_headers(ws, c1_mid_headers(cumulative_label))
    scope = wb[SHEET_C1_SCOPE]
    scope_hmap = header_map(scope)
    for row in range(2, scope.max_row + 1):
        out_row = row
        station = scope.cell(row, scope_hmap["站点"]).value
        main_sku = scope.cell(row, scope_hmap["主SKU分类"]).value
        operator = scope.cell(row, scope_hmap["运营"]).value
        if not station or not main_sku or not operator:
            continue
        for col, header in enumerate(C1_SCOPE_HEADERS, 1):
            ws.cell(out_row, col, scope.cell(row, scope_hmap[header]).value)
        ws.cell(
            out_row,
            6,
            f"=SUMIFS('{SHEET_C2_MID}'!$G:$G,"
            f"'{SHEET_C2_MID}'!$A:$A,$A{out_row},"
            f"'{SHEET_C2_MID}'!$F:$F,$D{out_row},"
            f"'{SHEET_C2_MID}'!$C:$C,$E{out_row})",
        )
    apply_formula_column_fill(ws, [6])
    return ws


def generate_middle_workbook(ctx, paths) :
    wb = Workbook()
    remove_default_sheet(wb)

    copy_mapping_sheet(paths["mapping"], SHEET_SKU_MAP, wb, SHEET_SKU_MAP)
    log_sheet_done(wb[SHEET_SKU_MAP])
    copy_mapping_sheet(paths["mapping"], SHEET_STORE_MAP, wb, SHEET_STORE_MAP)
    log_sheet_done(wb[SHEET_STORE_MAP])
    copy_mapping_sheet(paths["mapping"], SHEET_WAREHOUSE_MAP, wb, SHEET_WAREHOUSE_MAP)
    log_sheet_done(wb[SHEET_WAREHOUSE_MAP])

    sales_ws = create_sales_sheet(wb, first_sheet(paths["sales"]))
    log_sheet_done(sales_ws)
    sales_cum_ws = create_sales_sheet(wb, first_sheet(paths["sales_cum"]), SHEET_SALES_CUM)
    log_sheet_done(sales_cum_ws)
    receiving_ws = create_in_receiving_sheet(wb, first_sheet(paths["in_receiving"]))
    log_sheet_done(receiving_ws)
    log_sheet_done(create_feishu_sheet(wb, first_sheet(paths["feishu"]), ctx.sales_date))
    in_new_ws = create_in_new_sheet(wb, first_sheet(paths["in_new"]))
    log_sheet_done(in_new_ws)
    log_sheet_done(create_stock_sheet(wb, first_sheet(paths["stock"]), receiving_ws, in_new_ws))
    granularity_wb = load_workbook(paths["granularity"], data_only=True, read_only=False)
    try:
        scope_sources = [
            (SHEET_SKU_SCOPE, GRANULARITY_SKU, SKU_SCOPE_HEADERS),
            (SHEET_CAT_SCOPE, GRANULARITY_CAT, CAT_SCOPE_HEADERS),
            (SHEET_OP_SCOPE, GRANULARITY_OP, OP_SCOPE_HEADERS),
            (SHEET_C1_SCOPE, GRANULARITY_C1, C1_SCOPE_HEADERS),
            (SHEET_C2_SCOPE, GRANULARITY_C2, C2_SCOPE_HEADERS),
        ]
        for target_sheet, source_sheet, headers in scope_sources:
            scope_ws = ensure_sheet(wb, target_sheet)
            copy_source_to_headers(
                granularity_wb[source_sheet], scope_ws, headers, header_row=2, data_start_row=3
            )
            log_sheet_done(scope_ws, f"{target_sheet}（来源：{source_sheet}）")
    finally:
        granularity_wb.close()
    log_sheet_done(create_sku_mid_sheet(wb, ctx.sales_label))
    log_sheet_done(create_cat_mid_sheet(wb, ctx.sales_label))
    log_sheet_done(create_op_mid_sheet(wb, ctx.sales_label))
    log_sheet_done(create_c2_mid_sheet(wb, ctx.cumulative_label))
    log_sheet_done(create_c1_mid_sheet(wb, ctx.cumulative_label))

    order = [
        SHEET_SKU_MID, SHEET_CAT_MID, SHEET_OP_MID, SHEET_C1_MID, SHEET_C2_MID,
        SHEET_SKU_SCOPE, SHEET_CAT_SCOPE, SHEET_OP_SCOPE, SHEET_C1_SCOPE, SHEET_C2_SCOPE,
        SHEET_SALES, SHEET_SALES_CUM, SHEET_STOCK, SHEET_IN_NEW, SHEET_IN_RECEIVING,
        SHEET_FEISHU, SHEET_SKU_MAP, SHEET_STORE_MAP, SHEET_WAREHOUSE_MAP,
    ]
    wb._sheets = [wb[name] for name in order if name in wb.sheetnames]
    wb.save(ctx.middle_path)
    log_stage(f"已保存中间表：{ctx.middle_path}")


def ensure_result_workbook(path) :
    if path.exists():
        return load_workbook(path, data_only=False)
    wb = Workbook()
    remove_default_sheet(wb)
    wb.create_sheet(RESULT_SKU)
    wb.create_sheet(RESULT_CAT)
    wb.create_sheet(RESULT_OP)
    wb.create_sheet(RESULT_C1)
    wb.create_sheet(RESULT_C2)
    return wb


def ensure_result_sheet(wb, name) :
    if name in wb.sheetnames:
        return wb[name]
    return wb.create_sheet(name)


def ensure_headers(ws, required_prefix, extra_headers) :
    existing = [ws.cell(1, col).value for col in range(1, ws.max_column + 1)]
    if not any(existing):
        headers = required_prefix + extra_headers
        write_headers(ws, headers)
        return
    for col, header in enumerate(required_prefix, 1):
        ws.cell(1, col, header)
        ws.cell(1, col).font = Font(bold=True)
    hmap = header_map(ws)
    for header in extra_headers:
        if header not in hmap:
            ws.cell(1, ws.max_column + 1, header)


def ensure_base_headers(ws, required_prefix) :
    existing = [ws.cell(1, col).value for col in range(1, ws.max_column + 1)]
    if not any(existing):
        write_headers(ws, required_prefix)
        return
    for col, header in enumerate(required_prefix, 1):
        cell = ws.cell(1, col, header)
        cell.font = Font(bold=True)


def date_prefix(header, suffix) :
    if not isinstance(header, str):
        return None
    match = re.fullmatch(r"(\d{4}\.\d{2}\.\d{2})" + re.escape(suffix), header)
    return match.group(1) if match else None


def ensure_leftmost_date_column(ws, header, first_date_col) :
    hmap = header_map(ws)
    if header in hmap:
        return hmap[header]
    style_source_col = first_date_col if first_date_col <= ws.max_column else max(first_date_col - 1, 1)
    ws.insert_cols(first_date_col)
    copy_column_format(ws, style_source_col + 1 if style_source_col >= first_date_col else style_source_col, first_date_col)
    cell = ws.cell(1, first_date_col, header)
    return first_date_col


def copy_column_format(ws, source_col, target_col, max_row = None) :
    if source_col < 1 or target_col < 1 or source_col > ws.max_column:
        return
    limit = max_row or ws.max_row
    for row in range(1, limit + 1):
        copy_cell_style(ws.cell(row, source_col), ws.cell(row, target_col))
    source_letter = col_letter(source_col)
    target_letter = col_letter(target_col)
    ws.column_dimensions[target_letter].width = ws.column_dimensions[source_letter].width


def remove_category_growth_columns(ws) :
    for col in range(ws.max_column, 0, -1):
        if date_prefix(ws.cell(1, col).value, "涨幅"):
            ws.delete_cols(col)


def ensure_category_date_columns(ws, sales_label, growth_label) :
    first_date_col = len(CAT_RESULT_BASE_HEADERS) + 1
    hmap = header_map(ws)
    if sales_label not in hmap:
        style_source_col = first_date_col if first_date_col <= ws.max_column else max(first_date_col - 1, 1)
        ws.insert_cols(first_date_col, 2)
        source_after_insert = style_source_col + 2 if style_source_col >= first_date_col else style_source_col
        copy_column_format(ws, source_after_insert, first_date_col)
        copy_column_format(ws, first_date_col, first_date_col + 1)
        ws.cell(1, first_date_col, sales_label)
        ws.cell(1, first_date_col + 1, growth_label)
        return first_date_col, first_date_col + 1
    sales_col = hmap[sales_label]
    ws.insert_cols(sales_col + 1)
    copy_column_format(ws, sales_col, sales_col + 1)
    ws.cell(1, sales_col + 1, growth_label)
    return sales_col, sales_col + 1


def ensure_category_growth_columns(ws) :
    col = 1
    while col <= ws.max_column:
        header = ws.cell(1, col).value
        prefix = date_prefix(header, "销量")
        if prefix:
            growth_header = prefix + "涨幅"
            headers = header_map(ws)
            if growth_header not in headers:
                ws.insert_cols(col + 1)
                copy_column_format(ws, col, col + 1)
                ws.cell(1, col + 1, growth_header)
                col += 1
        col += 1


def invalid_result_key_value(value) :
    if value in (None, ""):
        return True
    if isinstance(value, str):
        text = value.strip()
        if text.startswith("="):
            return True
        return any(marker in text for marker in ["#REF", "#N/A", "#VALUE!", "#NAME?"])
    return False


def valid_source_key_value(value) :
    return not invalid_result_key_value(value)


def remove_invalid_result_rows(ws, key_headers) :
    hmap = header_map(ws)
    key_cols = [hmap[header] for header in key_headers if header in hmap]
    if len(key_cols) != len(key_headers):
        return
    for row in range(ws.max_row, 1, -1):
        key_values = [ws.cell(row, col).value for col in key_cols]
        if any(invalid_result_key_value(value) for value in key_values):
            if row_has_any_value(ws, row):
                ws.delete_rows(row)


def build_key_index(ws, key_headers) :
    hmap = header_map(ws)
    index = {}
    for row in range(2, ws.max_row + 1):
        key = canonical_key(ws.cell(row, hmap[h]).value for h in key_headers if h in hmap)
        if len(key) == len(key_headers) and all(valid_source_key_value(v) for v in key):
            index[key] = row
    return index


def granularity_configs() :
    return [
        {
            "result_sheet": RESULT_SKU,
            "source_sheet": SHEET_SKU_SCOPE,
            "result_keys": ["站点", "SKU"],
            "source_keys": ["站点(英文)", "产品编码"],
            "dimensions": [
                ("站点", "站点(英文)"),
                ("产品分级", "产品分级"),
                ("产品主类/材质", "产品主类/材质"),
                ("主SKU分类", "主SKU分类"),
                ("SKU", "产品编码"),
                ("产品名称", "产品名称"),
            ],
            "base_headers": SKU_RESULT_BASE_HEADERS,
            "source_header_row": 1,
            "source_data_start_row": 2,
        },
        {
            "result_sheet": RESULT_CAT,
            "source_sheet": SHEET_CAT_SCOPE,
            "result_keys": ["站点", "主SKU分类"],
            "source_keys": ["站点(英文)", "主SKU分类"],
            "dimensions": [
                ("站点", "站点(英文)"),
                ("产品分级", "产品分级"),
                ("产品主类/材质", "产品主类/材质"),
                ("主SKU分类", "主SKU分类"),
            ],
            "base_headers": CAT_RESULT_BASE_HEADERS,
            "source_header_row": 1,
            "source_data_start_row": 2,
        },
        {
            "result_sheet": RESULT_OP,
            "source_sheet": SHEET_OP_SCOPE,
            "result_keys": ["站点", "运营", "主SKU分类"],
            "source_keys": ["站点(英文)", "运营", "主SKU分类"],
            "dimensions": [
                ("站点", "站点(英文)"),
                ("产品分级", "产品分级"),
                ("产品主类/材质", "产品主类/材质"),
                ("主SKU分类", "主SKU分类"),
                ("运营", "运营"),
            ],
            "base_headers": OP_RESULT_BASE_HEADERS,
            "source_header_row": 1,
            "source_data_start_row": 2,
        },
        {
            "result_sheet": RESULT_C1,
            "source_sheet": SHEET_C1_SCOPE,
            "result_keys": ["站点", "运营", "主SKU分类"],
            "source_keys": ["站点", "运营", "主SKU分类"],
            "dimensions": [
                ("站点", "站点"),
                ("产品分级", "产品分级"),
                ("产品主类/材质", "产品主类/材质"),
                ("主SKU分类", "主SKU分类"),
                ("运营", "运营"),
            ],
            "base_headers": C1_RESULT_BASE_HEADERS,
            "source_header_row": 1,
            "source_data_start_row": 2,
        },
        {
            "result_sheet": RESULT_C2,
            "source_sheet": SHEET_C2_SCOPE,
            "result_keys": ["站点", "店铺名", "运营", "主SKU分类"],
            "source_keys": ["站点", "店铺名", "运营", "主SKU分类"],
            "dimensions": [
                ("站点", "站点"),
                ("店铺名", "店铺名"),
                ("运营", "运营"),
                ("产品分级", "产品分级"),
                ("产品主类/材质", "产品主类/材质"),
                ("主SKU分类", "主SKU分类"),
            ],
            "base_headers": C2_RESULT_BASE_HEADERS,
            "source_header_row": 1,
            "source_data_start_row": 2,
        },
    ]


def canonical_key(values) :
    return tuple(value.strip() if isinstance(value, str) else value for value in values)


def read_management_granularity(source_ws, config) :
    header_row = config.get("source_header_row", 2)
    data_start_row = config.get("source_data_start_row", header_row + 1)
    header_values = next(source_ws.iter_rows(min_row=header_row, max_row=header_row, values_only=True), None)
    if header_values is None:
        raise ValueError(f"粒度来源 {source_ws.title} 缺少第 {header_row} 行字段表头")
    source_header = {
        normalize_header(value): index
        for index, value in enumerate(header_values)
        if normalize_header(value)
    }
    missing_headers = [
        header
        for header in config["source_keys"] + [source for _, source in config["dimensions"]]
        if normalize_header(header) not in source_header
    ]
    if missing_headers:
        raise KeyError(f"粒度来源 {source_ws.title} 缺少字段：{sorted(set(missing_headers))}")
    source_key_cols = [source_header[normalize_header(header)] for header in config["source_keys"]]
    dimension_cols = {
        result_header: source_header[normalize_header(source_name)]
        for result_header, source_name in config["dimensions"]
    }
    rows = []
    seen = set()
    row_iter = source_ws.iter_rows(min_row=data_start_row, values_only=True)
    for row_number, row_values in enumerate(row_iter, start=data_start_row):
        raw_key = [row_values[col] if col < len(row_values) else None for col in source_key_cols]
        if all(value in (None, "") for value in raw_key):
            continue
        if any(invalid_result_key_value(value) for value in raw_key):
            raise ValueError(f"粒度来源 {source_ws.title} 第 {row_number} 行粒度字段无效：{raw_key}")
        key = canonical_key(raw_key)
        if key in seen:
            raise ValueError(f"粒度来源 {source_ws.title} 存在重复粒度：{key}")
        seen.add(key)
        dimensions = {}
        for result_header, col in dimension_cols.items():
            value = row_values[col] if col < len(row_values) else None
            if result_header in config["result_keys"] and isinstance(value, str):
                value = value.strip()
            dimensions[result_header] = value
        rows.append((key, dimensions))
    if not rows:
        raise ValueError(f"粒度来源 {source_ws.title} 没有有效粒度行")
    return rows


def snapshot_result_rows(ws, key_headers) :
    hmap = header_map(ws)
    key_cols = [hmap[header] for header in key_headers]
    order = []
    snapshots = {}
    invalid_rows = 0
    for row in range(2, ws.max_row + 1):
        if not row_has_any_value(ws, row):
            continue
        raw_key = [ws.cell(row, col).value for col in key_cols]
        if any(invalid_result_key_value(value) for value in raw_key):
            invalid_rows += 1
            continue
        key = canonical_key(raw_key)
        if key in snapshots:
            raise ValueError(f"结果表 {ws.title} 存在重复粒度，无法安全迁移历史数据：{key}")
        order.append(key)
        snapshots[key] = {
            "source_row": row,
            "values": [ws.cell(row, col).value for col in range(1, ws.max_column + 1)],
            "styles": [copy(ws.cell(row, col)._style) for col in range(1, ws.max_column + 1)],
            "comments": [copy(ws.cell(row, col).comment) for col in range(1, ws.max_column + 1)],
            "hyperlinks": [copy(ws.cell(row, col).hyperlink) for col in range(1, ws.max_column + 1)],
            "height": ws.row_dimensions[row].height,
            "hidden": ws.row_dimensions[row].hidden,
        }
    return order, snapshots, invalid_rows


def read_result_key_order(ws, key_headers) :
    hmap = header_map(ws)
    key_cols = [hmap[header] for header in key_headers]
    order = []
    seen = set()
    invalid_rows = 0
    for row in range(2, ws.max_row + 1):
        if not row_has_any_value(ws, row):
            continue
        raw_key = [ws.cell(row, col).value for col in key_cols]
        if any(invalid_result_key_value(value) for value in raw_key):
            invalid_rows += 1
            continue
        key = canonical_key(raw_key)
        if key in seen:
            raise ValueError(f"结果表 {ws.title} 存在重复粒度，无法安全迁移历史数据：{key}")
        seen.add(key)
        order.append(key)
    return order, invalid_rows


def comparable_dimension_value(value) :
    if value in (None, ""):
        return None
    return value.strip() if isinstance(value, str) else value


def find_dimension_differences(
    ws,
    management_rows,
) :
    hmap = header_map(ws)
    differences = []
    for target_row, (_, dimensions) in enumerate(management_rows, start=2):
        for header, expected_value in dimensions.items():
            actual_value = ws.cell(target_row, hmap[header]).value
            if comparable_dimension_value(actual_value) != comparable_dimension_value(expected_value):
                differences.append((target_row, header, expected_value))
    return differences


def write_snapshot_row(ws, target_row, snapshot, template) :
    style_source = snapshot or template
    source_row = snapshot["source_row"] if snapshot else None
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(target_row, col)
        value = snapshot["values"][col - 1] if snapshot else None
        if source_row and isinstance(value, str) and value.startswith("=") and source_row != target_row:
            origin = f"{col_letter(col)}{source_row}"
            target = f"{col_letter(col)}{target_row}"
            try:
                value = Translator(value, origin=origin).translate_formula(target)
            except Exception:
                pass
        cell.value = value
        if style_source:
            cell._style = copy(style_source["styles"][col - 1])
            cell.comment = copy(style_source["comments"][col - 1]) if snapshot else None
            cell.hyperlink = copy(style_source["hyperlinks"][col - 1]) if snapshot else None
    if style_source:
        ws.row_dimensions[target_row].height = style_source["height"]
        ws.row_dimensions[target_row].hidden = style_source["hidden"]


def align_result_sheet_to_management(ws, source_ws, config) :
    ensure_base_headers(ws, config["base_headers"])
    management_rows = read_management_granularity(source_ws, config)
    expected_order = [key for key, _ in management_rows]
    current_order, invalid_rows = read_result_key_order(ws, config["result_keys"])
    if invalid_rows == 0 and current_order == expected_order:
        differences = find_dimension_differences(ws, management_rows)
        hmap = header_map(ws)
        for row, header, value in differences:
            ws.cell(row, hmap[header], value)
        if differences:
            print(
                f"[粒度对齐] {ws.title}：粒度与排序一致，已完成全量维度比对并刷新；"
                f"行数={len(expected_order)}，维度更新单元格={len(differences)}"
            )
        else:
            print(
                f"[粒度对齐] {ws.title}：已完成全量比对，粒度、排序及维度数据均一致；"
                f"行数={len(expected_order)}"
            )
        return

    _, snapshots, _ = snapshot_result_rows(ws, config["result_keys"])

    current_keys = set(snapshots)
    expected_keys = set(expected_order)
    added = len(expected_keys - current_keys)
    removed = len(current_keys - expected_keys) + invalid_rows
    moved = sum(1 for idx, key in enumerate(expected_order) if idx >= len(current_order) or current_order[idx] != key)
    template = next(iter(snapshots.values()), None)
    hmap = header_map(ws)

    for offset, (key, dimensions) in enumerate(management_rows, start=2):
        write_snapshot_row(ws, offset, snapshots.get(key), template)
        for header, value in dimensions.items():
            ws.cell(offset, hmap[header], value)

    target_last_row = len(management_rows) + 1
    if ws.max_row > target_last_row:
        ws.delete_rows(target_last_row + 1, ws.max_row - target_last_row)
    print(
        f"[粒度对齐] {ws.title}：已按 {source_ws.title} 重建顺序；"
        f"目标行数={len(expected_order)}，新增={added}，移除={removed}，位置变化={moved}"
    )


def align_result_granularity(wb, middle_path) :
    source_wb = load_workbook(middle_path, read_only=True, data_only=True)
    try:
        for config in granularity_configs():
            ws = ensure_result_sheet(wb, config["result_sheet"])
            source_ws = source_wb[config["source_sheet"]]
            align_result_sheet_to_management(ws, source_ws, config)
    finally:
        source_wb.close()
    log_stage(f"已按每日数据中间表数据口径完成结果表粒度与排序对齐：{middle_path.name}")


def validate_result_granularity_alignment(result_wb, middle_path) :
    source_wb = load_workbook(middle_path, read_only=True, data_only=True)
    try:
        for config in granularity_configs():
            management_rows = read_management_granularity(source_wb[config["source_sheet"]], config)
            expected = [key for key, _ in management_rows]
            actual, invalid_rows = read_result_key_order(result_wb[config["result_sheet"]], config["result_keys"])
            if invalid_rows or actual != expected:
                raise RuntimeError(
                    f"结果表 {config['result_sheet']} 的粒度或排序与 {config['source_sheet']} 不一致"
                )
            differences = find_dimension_differences(
                result_wb[config["result_sheet"]],
                management_rows,
            )
            if differences:
                raise RuntimeError(
                    f"结果表 {config['result_sheet']} 仍有 {len(differences)} 个维度单元格与中间表数据口径不一致"
                )
    finally:
        source_wb.close()


def middle_value(formula_ws, values_ws, row, col) :
    """读取中间表静态值；遇到公式时只返回 Excel 缓存值，并保留 #N/A 等错误结果。"""
    value = formula_ws.cell(row, col).value
    if isinstance(value, str) and value.startswith("="):
        cached = values_ws.cell(row, col).value
        return cached if cached not in (None, "") else None
    return value


def reset_sales_column(ws, sales_col) :
    for row in range(2, ws.max_row + 1):
        if row_has_any_value(ws, row):
            ws.cell(row, sales_col, 0)


def clear_data_columns_before_inventory_status(ws, data_headers) :
    hmap = header_map(ws)
    if "库存状况" not in hmap:
        raise KeyError(f"结果表 {ws.title} 缺少字段：库存状况")
    status_col = hmap["库存状况"]
    clear_headers = [
        header
        for header in data_headers
        if header != "库存状况" and header in hmap and hmap[header] < status_col
    ]
    for row in range(2, ws.max_row + 1):
        for header in clear_headers:
            ws.cell(row, hmap[header]).value = None
    print(
        f"[数据清空] {ws.title}：已清空本次指定的数据刷新列；"
        f"字段数={len(clear_headers)}，数据行={max(ws.max_row - 1, 0)}"
    )


def validate_middle_formula_cache(ctx) :
    wb = load_workbook(ctx.middle_path, data_only=True, read_only=True)
    checks = [
        (SHEET_SKU_MID, ctx.sales_label),
        (SHEET_CAT_MID, ctx.sales_label),
        (SHEET_OP_MID, ctx.sales_label),
        (SHEET_C1_MID, ctx.cumulative_label),
        (SHEET_C2_MID, ctx.cumulative_label),
    ]
    for sheet_name, header in checks:
        ws = wb[sheet_name]
        headers = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
        if header not in headers:
            raise RuntimeError(f"{sheet_name} 缺少公式缓存校验字段：{header}")
        col_idx = headers.index(header)
        data_rows = 0
        cached_rows = 0
        for row in ws.iter_rows(min_row=2, values_only=True):
            if any(value not in (None, "") for value in row):
                data_rows += 1
                if row[col_idx] not in (None, ""):
                    cached_rows += 1
        if data_rows and cached_rows == 0:
            raise RuntimeError(f"{sheet_name} 的 {header} 公式未产生可读取缓存值，请确认 Excel 公式计算已成功")
        print(f"[公式缓存] {sheet_name} / {header}：数据行={data_rows}，可读取缓存行={cached_rows}")
    log_stage("已确认中间表公式缓存可读取")


def write_result_value(ws, row, col, value, allow_blank = False) :
    # 结果表写入的是中间表缓存结果，不复制中间表公式。
    if value is None and not allow_blank:
        return
    ws.cell(row, col, value)


def refresh_sku_result_formulas(ws) :
    hmap = header_map(ws)
    factory_col = col_letter(hmap["货值单价"])
    landed_col = col_letter(hmap["货值+海运单价"])
    sales_col = col_letter(hmap["销售额单价"])
    billing_col = col_letter(hmap["账单收入单价"])
    transit_groups = [
        ("30天内到货在途数量", "30天货值", "30天货值+海运", "30天销售额", "30天账单收入"),
        ("60天内到货在途数量", "60天货值", "60天货值+海运", "60天销售额", "60天账单收入"),
        ("90天在途数量", "90天货值", "90天货值+海运", "90天销售额", "90天账单收入"),
    ]

    for row in range(2, ws.max_row + 1):
        if not valid_source_key_value(ws.cell(row, hmap["SKU"]).value):
            continue
        for quantity_header, value_header, landed_value_header, sales_value_header, billing_value_header in transit_groups:
            quantity_col = col_letter(hmap[quantity_header])
            ws.cell(row, hmap[value_header], f"={quantity_col}{row}*{factory_col}{row}")
            ws.cell(row, hmap[landed_value_header], f"={quantity_col}{row}*{landed_col}{row}")
            ws.cell(row, hmap[sales_value_header], f"={quantity_col}{row}*{sales_col}{row}")
            ws.cell(row, hmap[billing_value_header], f"={quantity_col}{row}*{billing_col}{row}")
    print(
        f"[公式刷新] {ws.title}：四个单价已从中间表同步，已刷新12个30/60/90天派生公式；"
        f"数据行={max(ws.max_row - 1, 0)}"
    )


def update_result_sku(ctx, wb) :
    ws = ensure_result_sheet(wb, RESULT_SKU)
    ensure_base_headers(ws, SKU_RESULT_BASE_HEADERS)
    remove_invalid_result_rows(ws, ["站点", "SKU"])
    sales_col = ensure_leftmost_date_column(ws, ctx.sales_label, len(SKU_RESULT_BASE_HEADERS) + 1)
    clear_data_columns_before_inventory_status(ws, SKU_RESULT_CLEAR_HEADERS)
    reset_sales_column(ws, sales_col)
    hmap = header_map(ws)
    source_wb = load_workbook(ctx.middle_path, data_only=False)
    source_values_wb = load_workbook(ctx.middle_path, data_only=True)
    source = source_wb[SHEET_SKU_MID]
    source_values = source_values_wb[SHEET_SKU_MID]
    src_hmap = header_map(source)
    key_index = build_key_index(ws, ["站点", "SKU"])
    skipped_unmanaged = 0

    for src_row in range(2, source.max_row + 1):
        station = middle_value(source, source_values, src_row, src_hmap["站点"])
        sku = middle_value(source, source_values, src_row, src_hmap["SKU"])
        if not valid_source_key_value(station) or not valid_source_key_value(sku):
            continue
        key = canonical_key((station, sku))
        dst_row = key_index.get(key)
        if dst_row is None:
            skipped_unmanaged += 1
            continue
        for header in SKU_RESULT_DATA_HEADERS:
            if header not in hmap or header not in src_hmap:
                continue
            dst_col = hmap[header]
            src_col = src_hmap[header]
            write_result_value(ws, dst_row, dst_col, middle_value(source, source_values, src_row, src_col), header in RESULT_BLANK_FIELDS)
        write_result_value(ws, dst_row, sales_col, middle_value(source, source_values, src_row, src_hmap[ctx.sales_label]))
    source_wb.close()
    source_values_wb.close()
    refresh_sku_result_formulas(ws)
    if skipped_unmanaged:
        print(f"[粒度过滤] {RESULT_SKU}：指标表存在但数据口径表不存在，已跳过={skipped_unmanaged}")
    log_sheet_done(ws, RESULT_SKU)


def refresh_category_result_formulas(ws, sku_ws) :
    hmap = header_map(ws)
    sku_hmap = header_map(sku_ws)
    station_col = col_letter(hmap["站点"])
    main_sku_col = col_letter(hmap["主SKU分类"])
    sku_station_col = col_letter(sku_hmap["站点"])
    sku_main_sku_col = col_letter(sku_hmap["主SKU分类"])
    sku_sheet = f"'{RESULT_SKU}'"

    for header in CAT_RESULT_SUM_FORMULA_HEADERS:
        if header not in sku_hmap:
            raise KeyError(f"结果表 {RESULT_SKU} 缺少类目汇总字段：{header}")

    sales_headers = [header for header in hmap if date_prefix(header, "销量")]
    missing_sales_headers = [header for header in sales_headers if header not in sku_hmap]
    if missing_sales_headers:
        raise KeyError(f"结果表 {RESULT_SKU} 缺少类目日销所需历史销量字段：{missing_sales_headers}")

    for row in range(2, ws.max_row + 1):
        if not valid_source_key_value(ws.cell(row, hmap["主SKU分类"]).value):
            continue
        for header in CAT_RESULT_SUM_FORMULA_HEADERS:
            sku_value_col = col_letter(sku_hmap[header])
            ws.cell(
                row,
                hmap[header],
                f"=SUMIFS({sku_sheet}!{sku_value_col}:{sku_value_col},"
                f"{sku_sheet}!${sku_station_col}:${sku_station_col},${station_col}{row},"
                f"{sku_sheet}!${sku_main_sku_col}:${sku_main_sku_col},${main_sku_col}{row})",
            )
        actual_stock_col = col_letter(hmap["当日库存"])
        monthly_sales_col = col_letter(hmap["月销量"])
        ws.cell(row, hmap["可售月数"], f"=IFERROR({actual_stock_col}{row}/{monthly_sales_col}{row},0)")
        ws.cell(row, hmap["库存状况"]).value = None
        for header in sales_headers:
            sku_sales_col = col_letter(sku_hmap[header])
            ws.cell(
                row,
                hmap[header],
                f"=SUMIFS({sku_sheet}!{sku_sales_col}:{sku_sales_col},"
                f"{sku_sheet}!${sku_station_col}:${sku_station_col},${station_col}{row},"
                f"{sku_sheet}!${sku_main_sku_col}:${sku_main_sku_col},${main_sku_col}{row})",
            )
    print(
        f"[公式刷新] {ws.title}：已按 {RESULT_SKU} 刷新类目指标和全部日期销量公式；"
        f"数据行={max(ws.max_row - 1, 0)}，销量列={len(sales_headers)}"
    )


def sync_category_supplier(ctx, ws) :
    hmap = header_map(ws)
    supplier_col = hmap["供应商"]
    for row in range(2, ws.max_row + 1):
        ws.cell(row, supplier_col).value = None

    source_wb = load_workbook(ctx.middle_path, data_only=False)
    source_values_wb = load_workbook(ctx.middle_path, data_only=True)
    try:
        source = source_wb[SHEET_CAT_MID]
        source_values = source_values_wb[SHEET_CAT_MID]
        src_hmap = header_map(source)
        key_index = build_key_index(ws, ["站点", "主SKU分类"])
        synced = 0
        for src_row in range(2, source.max_row + 1):
            station = middle_value(source, source_values, src_row, src_hmap["站点"])
            main_sku = middle_value(source, source_values, src_row, src_hmap["主SKU分类"])
            if not valid_source_key_value(station) or not valid_source_key_value(main_sku):
                continue
            dst_row = key_index.get(canonical_key((station, main_sku)))
            if dst_row is None:
                continue
            supplier = middle_value(source, source_values, src_row, src_hmap["供应商"])
            write_result_value(ws, dst_row, supplier_col, supplier, allow_blank=True)
            synced += 1
        print(f"[字段同步] {RESULT_CAT} / 供应商：已按站点+主SKU分类同步={synced}")
    finally:
        source_wb.close()
        source_values_wb.close()


def update_result_cat(ctx, wb) :
    ws = ensure_result_sheet(wb, RESULT_CAT)
    ensure_base_headers(ws, CAT_RESULT_BASE_HEADERS)
    remove_invalid_result_rows(ws, ["站点", "主SKU分类"])
    remove_category_growth_columns(ws)
    ensure_category_date_columns(ws, ctx.sales_label, ctx.growth_label)
    ensure_category_growth_columns(ws)
    sync_category_supplier(ctx, ws)
    refresh_category_result_formulas(ws, wb[RESULT_SKU])
    refresh_category_growth_formulas(ws)
    log_sheet_done(ws, RESULT_CAT)


def refresh_category_growth_formulas(ws) :
    hmap = header_map(ws)
    sales_headers = []
    for header, col in hmap.items():
        prefix = date_prefix(header, "销量")
        if prefix:
            sales_headers.append((prefix, col, header))
    sales_headers.sort(reverse=True)
    for idx, (prefix, sales_col, _) in enumerate(sales_headers):
        growth_header = prefix + "涨幅"
        if growth_header not in hmap:
            continue
        growth_col = hmap[growth_header]
        previous_col = sales_headers[idx + 1][1] if idx + 1 < len(sales_headers) else None
        for row in range(2, ws.max_row + 1):
            if previous_col:
                cur = f"{col_letter(sales_col)}{row}"
                prev = f"{col_letter(previous_col)}{row}"
                ws.cell(row, growth_col, f"=IFERROR(({cur}-{prev})/{cur},0)")
            else:
                ws.cell(row, growth_col, None)


def update_result_op(ctx, wb) :
    ws = ensure_result_sheet(wb, RESULT_OP)
    ensure_base_headers(ws, OP_RESULT_BASE_HEADERS)
    remove_invalid_result_rows(ws, ["站点", "运营", "主SKU分类"])
    sales_col = ensure_leftmost_date_column(ws, ctx.sales_label, len(OP_RESULT_BASE_HEADERS) + 1)
    reset_sales_column(ws, sales_col)
    hmap = header_map(ws)
    source_wb = load_workbook(ctx.middle_path, data_only=False)
    source_values_wb = load_workbook(ctx.middle_path, data_only=True)
    source = source_wb[SHEET_OP_MID]
    source_values = source_values_wb[SHEET_OP_MID]
    src_hmap = header_map(source)
    key_index = build_key_index(ws, ["站点", "运营", "主SKU分类"])
    skipped_unmanaged = 0

    for src_row in range(2, source.max_row + 1):
        station = middle_value(source, source_values, src_row, src_hmap["站点"])
        operator = middle_value(source, source_values, src_row, src_hmap["运营"])
        main_sku = middle_value(source, source_values, src_row, src_hmap["主SKU分类"])
        if not valid_source_key_value(station) or not valid_source_key_value(operator) or not valid_source_key_value(main_sku):
            continue
        key = canonical_key((station, operator, main_sku))
        dst_row = key_index.get(key)
        if dst_row is None:
            skipped_unmanaged += 1
            continue
        for header in OP_RESULT_DATA_HEADERS:
            if header in hmap and header in src_hmap:
                write_result_value(ws, dst_row, hmap[header], middle_value(source, source_values, src_row, src_hmap[header]), header in RESULT_BLANK_FIELDS)
        write_result_value(ws, dst_row, sales_col, middle_value(source, source_values, src_row, src_hmap[ctx.sales_label]))
    source_wb.close()
    source_values_wb.close()
    if skipped_unmanaged:
        print(f"[粒度过滤] {RESULT_OP}：指标表存在但数据口径表不存在，已跳过={skipped_unmanaged}")
    log_sheet_done(ws, RESULT_OP)


def update_result_cumulative(
    ctx,
    wb,
    result_sheet,
    middle_sheet,
    base_headers,
    key_headers,
) :
    ws = ensure_result_sheet(wb, result_sheet)
    ensure_base_headers(ws, base_headers)
    remove_invalid_result_rows(ws, key_headers)
    cumulative_col = ensure_leftmost_date_column(ws, ctx.cumulative_label, len(base_headers) + 1)
    reset_sales_column(ws, cumulative_col)
    hmap = header_map(ws)
    source_wb = load_workbook(ctx.middle_path, data_only=False)
    source_values_wb = load_workbook(ctx.middle_path, data_only=True)
    try:
        source = source_wb[middle_sheet]
        source_values = source_values_wb[middle_sheet]
        src_hmap = header_map(source)
        key_index = build_key_index(ws, key_headers)
        synced = 0
        skipped_unmanaged = 0
        for src_row in range(2, source.max_row + 1):
            raw_key = [middle_value(source, source_values, src_row, src_hmap[header]) for header in key_headers]
            if any(not valid_source_key_value(value) for value in raw_key):
                continue
            dst_row = key_index.get(canonical_key(raw_key))
            if dst_row is None:
                skipped_unmanaged += 1
                continue
            write_result_value(
                ws,
                dst_row,
                cumulative_col,
                middle_value(source, source_values, src_row, src_hmap[ctx.cumulative_label]),
            )
            synced += 1
        if skipped_unmanaged:
            print(f"[粒度过滤] {result_sheet}：指标表存在但数据口径表不存在，已跳过={skipped_unmanaged}")
        print(f"[字段同步] {result_sheet} / {ctx.cumulative_label}：已同步={synced}")
    finally:
        source_wb.close()
        source_values_wb.close()
    log_sheet_done(ws, result_sheet)


def generate_result_workbook(ctx) :
    wb = ensure_result_workbook(ctx.result_path)
    align_result_granularity(wb, ctx.middle_path)
    update_result_sku(ctx, wb)
    update_result_cat(ctx, wb)
    update_result_op(ctx, wb)
    update_result_cumulative(ctx, wb, RESULT_C1, SHEET_C1_MID, C1_RESULT_BASE_HEADERS, ["站点", "运营", "主SKU分类"])
    update_result_cumulative(ctx, wb, RESULT_C2, SHEET_C2_MID, C2_RESULT_BASE_HEADERS, ["站点", "店铺名", "运营", "主SKU分类"])
    wb.save(ctx.result_path)
    log_stage(f"已保存结果表：{ctx.result_path}")


def validate_output(ctx) :
    middle_wb = load_workbook(ctx.middle_path, data_only=False, read_only=True)
    required_middle_sheets = {
        SHEET_SKU_MID, SHEET_CAT_MID, SHEET_OP_MID, SHEET_C1_MID, SHEET_C2_MID,
        SHEET_SKU_SCOPE, SHEET_CAT_SCOPE, SHEET_OP_SCOPE, SHEET_C1_SCOPE, SHEET_C2_SCOPE,
        SHEET_SALES, SHEET_SALES_CUM, SHEET_STOCK, SHEET_IN_NEW, SHEET_IN_RECEIVING,
        SHEET_FEISHU, SHEET_SKU_MAP, SHEET_STORE_MAP, SHEET_WAREHOUSE_MAP,
    }
    missing_middle = required_middle_sheets - set(middle_wb.sheetnames)
    if missing_middle:
        raise RuntimeError(f"中间表缺少 sheet：{sorted(missing_middle)}")
    if ctx.sales_label not in header_map(middle_wb[SHEET_SKU_MID]):
        raise RuntimeError(f"{SHEET_SKU_MID} 缺少销量日期字段：{ctx.sales_label}")
    for sheet in [SHEET_C1_MID, SHEET_C2_MID]:
        if ctx.cumulative_label not in header_map(middle_wb[sheet]):
            raise RuntimeError(f"{sheet} 缺少累计销量日期字段：{ctx.cumulative_label}")

    # read_only 模式下 ws.cell() 反复随机访问会极慢；结果表体量较小，校验使用普通模式读取。
    result_wb = load_workbook(ctx.result_path, data_only=False, read_only=False)
    for sheet in [RESULT_SKU, RESULT_CAT, RESULT_OP, RESULT_C1, RESULT_C2]:
        if sheet not in result_wb.sheetnames:
            raise RuntimeError(f"结果表缺少 sheet：{sheet}")
    if ctx.sales_label not in header_map(result_wb[RESULT_SKU]):
        raise RuntimeError(f"{RESULT_SKU} 缺少销量日期字段：{ctx.sales_label}")
    if ctx.growth_label not in header_map(result_wb[RESULT_CAT]):
        raise RuntimeError(f"{RESULT_CAT} 缺少涨幅日期字段：{ctx.growth_label}")
    for sheet in [RESULT_C1, RESULT_C2]:
        if ctx.cumulative_label not in header_map(result_wb[sheet]):
            raise RuntimeError(f"{sheet} 缺少累计销量日期字段：{ctx.cumulative_label}")
    result_key_headers = {
        RESULT_SKU: ["站点", "SKU"],
        RESULT_CAT: ["站点", "主SKU分类"],
        RESULT_OP: ["站点", "运营", "主SKU分类"],
        RESULT_C1: ["站点", "运营", "主SKU分类"],
        RESULT_C2: ["站点", "店铺名", "运营", "主SKU分类"],
    }
    for sheet, key_headers in result_key_headers.items():
        ws = result_wb[sheet]
        hmap = header_map(ws)
        for row in range(2, ws.max_row + 1):
            for header in key_headers:
                value = ws.cell(row, hmap[header]).value
                if invalid_result_key_value(value):
                    raise RuntimeError(f"{sheet} 第 {row} 行关键粒度字段无效：{header}={value}")
            for col in range(1, ws.max_column + 1):
                value = ws.cell(row, col).value
                if isinstance(value, str) and value.startswith("="):
                    if "#REF" in value or ("[" in value and "]" in value):
                        raise RuntimeError(f"{sheet} 第 {row} 行第 {col} 列存在不稳定公式：{value}")
            product_material = ws.cell(row, hmap["产品主类/材质"]).value
            if isinstance(product_material, str) and product_material.startswith("="):
                raise RuntimeError(f"{sheet} 第 {row} 行产品主类/材质必须为静态缓存值，不得保留公式")
            if sheet == RESULT_SKU:
                for header in SKU_RESULT_FORMULA_HEADERS:
                    formula = ws.cell(row, hmap[header]).value
                    if not isinstance(formula, str) or not formula.startswith("="):
                        raise RuntimeError(f"{sheet} 第 {row} 行 {header} 必须保留内部公式")
                for header in SKU_RESULT_PRICE_SYNC_HEADERS:
                    value = ws.cell(row, hmap[header]).value
                    if isinstance(value, str) and value.startswith("="):
                        raise RuntimeError(f"{sheet} 第 {row} 行 {header} 必须为中间表缓存静态值，不得保留结果表查询公式")
            if sheet == RESULT_CAT:
                formula_headers = set(CAT_RESULT_FORMULA_HEADERS) | {
                    header for header in hmap if date_prefix(header, "销量")
                }
                for header in formula_headers:
                    formula = ws.cell(row, hmap[header]).value
                    if not isinstance(formula, str) or not formula.startswith("="):
                        raise RuntimeError(f"{sheet} 第 {row} 行 {header} 必须保留引用 {RESULT_SKU} 的内部公式")
    validate_result_granularity_alignment(result_wb, ctx.middle_path)
    middle_wb.close()
    result_wb.close()


def run(root, date) :
    """按给定根目录和执行日期生成中间表并更新结果表。

    其他 Python 文件也可以 import 本函数后调用：
        from 生成每日销售及库存跟进表 import run
        run(r"<ROOT>", "2026-06-17")
    """
    ctx = build_context(root, date)
    try:
        paths = source_paths(ctx)
        validate_sources(paths)
        log_stage("已完成源文件存在性与有效数据校验")
        generate_middle_workbook(ctx, paths)
        recalculate_workbook_with_excel(ctx.middle_path)
        validate_middle_formula_cache(ctx)
        generate_result_workbook(ctx)
        recalculate_workbook_with_excel(ctx.result_path, workbook_label="结果表")
        validate_output(ctx)
        log_stage("已完成输出文件结构与关键字段校验")
        print(f"已生成中间表：{ctx.middle_path}")
        print(f"已更新结果表：{ctx.result_path}")
        return ctx
    finally:
        release_resources_and_verify([ctx.middle_path, ctx.result_path])


# def main(root = None, date = None) :
#     run_root, run_date = resolve_run_parameters(root, date)
#     run(run_root, run_date)
#     return 0


if __name__ == "__main__":
    try:
        raise SystemExit(run(r"<ROOT>", "2026-07-03"))
    except Exception as exc:
        print(f"执行失败：{exc}", file=sys.stderr)
        raise
