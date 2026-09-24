import argparse
import gc
import os
import re
import shutil
import subprocess
import tempfile
import time
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter
from copy import copy
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import from_excel


SHEET_ORDER = [
    "A1类目销售数据表", "A2类目SKU销售监控表", "B1运营类目销售表", "B2店铺类目销售表",
    "A1类目销售数据口径", "A2类目SKU销售监控数据口径", "B1运营类目数据口径", "B2店铺类目数据口径",
    "千易商品销售统计", "千易库存数据", "千易入库单_收货中",
    "SKU信息表", "店铺信息表", "仓库信息表", "净需求-目标表",
]

HEADERS = {
    "A1类目销售数据表": ["供应商", "站点", "产品分级", "产品主类/材质", "主SKU分类", "月份", "季度", "销量", "销售额", "销售额占比", "计划目标", "实际库存", "备货准确率", "周转系数"],
    "A2类目SKU销售监控表": ["供应商", "站点", "产品分级", "产品主类/材质", "主SKU分类", "产品编码", "产品名称", "月份", "季度", "计划目标", "实际销量", "目标完成率", "实际库存", "备货准确率", "周转系数"],
    "B1运营类目销售表": ["站点", "产品分级", "产品主类/材质", "主SKU分类", "运营", "月份", "季度", "实际销量", "实际销售额", "销售额占比"],
    "B2店铺类目销售表": ["站点", "店铺名", "运营", "产品分级", "产品主类/材质", "主SKU分类", "月份", "季度", "实际库存", "实际销量", "客单价（RMB）", "实际销售额"],
    "A1类目销售数据口径": ["供应商", "站点（中文）", "站点（英文）", "产品分级", "产品主类/材质", "主SKU分类"],
    "A2类目SKU销售监控数据口径": ["供应商", "站点（中文）", "站点（英文）", "产品分级", "产品主类/材质", "主SKU分类", "产品编码", "产品名称"],
    "B1运营类目数据口径": ["站点（中文）", "站点（英文）", "产品分级", "产品主类/材质", "类目", "运营"],
    "B2店铺类目数据口径": ["站点", "店铺名", "运营", "产品分级", "产品主类/材质", "主SKU分类"],
    "千易商品销售统计": ["月份", "季度", "系统sku", "sku名称", "spu 编码", "spu", "spu值", "主SKU分类", "币种", "平台", "店铺", "运营", "站点", "站点(中文)", "仓库", "日均销量", "销售额", "销售额VALUE", "销售订单量", "销售sku数量", "销售sku数量VALUE", "销售SKU数量(不限)", "发货前取消订单量", "发货前取消金额", "取消sku数量", "发货后退款订单量", "退款sku数量", "平均售价", "平均售价VALUE", "平均sku数量", "手工单量", "手工单sku数量", "补发货单量", "补发货sku数量", "时间类型", "时间范围", "日期Value"],
    "千易库存数据": ["月份", "季度", "商品编码", "商品名称", "主SKU分类", "仓库", "站点（中文）", "可用量", "在途量", "未收到库存", "实际在途量", "库存合计", "3天销量", "7天销量", "15天销量", "30天销量", "60天销量", "90天销量"],
    "千易入库单_收货中": ["入库单号", "参考号", "业务单号", "供应商", "采购员", "1688单号", "收货仓", "目的仓", "入库类型", "物流名称", "物流商", "跟踪号", "状态", "关联订单的出库批次", "商品图片", "商品编码", "商品名称", "单品尺寸-长", "单品尺寸-宽", "单品尺寸-高", "单品毛重", "单品净重", "自定义1", "自定义2", "商品备注", "采购费用", "头程费用", "调拨费用", "预计数量", "已收数量", "仓库收货数量", "仓库上架数量", "差异数据", "仓库收货明细", "仓库上架明细", " 已收良品量", "已收不良品量", "头程中转仓入库单号", "头程中转入库业务单号", "对应采购单参考号", "创建人", "创建时间", "收货人", "入库时间", "完成时间", "备注", "批注"],
    "店铺信息表": ["国家", "店铺", "运营", "站点"],
    "仓库信息表": ["仓库代码", "对应站点", "站点缩写"],
}

GRAIN_SOURCES = {
    "A1类目销售数据口径": "A1类目销售数据表（粒度）",
    "A2类目SKU销售监控数据口径": "A2类目SKU销售监控表（粒度）",
    "B1运营类目数据口径": "B1运营类目销售表（粒度）",
    "B2店铺类目数据口径": "B2店铺类目销售数据（粒度） ",
}

SITE_CODES = {"菲律宾": "PH", "马来": "MY", "泰国": "TH", "印尼1": "ID"}
HEADER_FILL = PatternFill("solid", fgColor="4472C4")
HEADER_FONT = Font(color="FFFFFF", bold=True)
FORMULA_FILL = PatternFill("solid", fgColor="DDEBF7")
THIN = Side(style="thin", color="D9E1F2")
CELL_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


class Context:
    """保存一次月度任务所需的运行上下文。"""

    def __init__(self, root, month, month_cn, quarter_cn, month_dot, output, paths):
        self.root = root
        self.month = month
        self.month_cn = month_cn
        self.quarter_cn = quarter_cn
        self.month_dot = month_dot
        self.output = output
        self.paths = paths


def log(message) :
    print(f"[{datetime.now():%H:%M:%S}] {message}", flush=True)


def verify_file_released(path) :
    if not path.exists():
        return
    probe = path.with_name(f".{path.name}.lockcheck.{os.getpid()}.tmp")
    try:
        path.replace(probe)
        probe.replace(path)
    except PermissionError as exc:
        raise PermissionError(f"文件仍被 Python、Excel 或其他程序占用，无法移动：{path}") from exc
    finally:
        if probe.exists() and not path.exists():
            probe.replace(path)


def release_resources_and_verify(paths) :
    gc.collect()
    time.sleep(0.5)
    for path in paths:
        verify_file_released(path)
    log("资源释放校验完成：输出 Excel 未被当前任务占用")


def parse_month(text) :
    if not re.fullmatch(r"\d{4}-\d{2}", text):
        raise ValueError("执行月份必须使用 YYYY-MM 格式，例如 2026-06")
    value = datetime.strptime(text, "%Y-%m").date().replace(day=1)
    return value


def add_months(value, offset) :
    index = value.year * 12 + value.month - 1 + offset
    return date(index // 12, index % 12 + 1, 1)


def build_context(root_text, month_text) :
    root = Path(root_text).expanduser().resolve()
    month = parse_month(month_text)
    month_cn = f"{month.year}年{month.month:02d}月"
    quarter_cn = f"{month.year}年第{(month.month - 1) // 3 + 1}季度"
    month_dot = f"{month.year}.{month.month:02d}"
    snapshot_day = month.strftime("%Y-%m-01")
    paths = {
        "sales": root / f"源数据/千易销售数据/商品销售统计_{month_text}.xlsx",
        "stock": root / f"源数据/千易库存数据/库存查询_{snapshot_day}.xlsx",
        "receiving": root / f"源数据/千易入库单数据/入库单明细_收货中_{snapshot_day}.xlsx",
        "mapping": root / "源数据/映射表/SKU、店铺、仓库映射表.xlsx",
        "grain": root / "源数据/映射表/@数据粒度表_运营大盘数据监控表.xlsx",
        "dashboard": root / "结果表/运营大盘数据监控表.xlsx",
    }
    for site in SITE_CODES:
        paths[f"demand_{site}"] = root / f"数据中间表/净需求计算表/净需求计算表_{site}_{month_cn}.xlsx"
    return Context(root, month, month_cn, quarter_cn, month_dot, root / f"数据中间表/月度数据中间表/月度数据中间表_{month_cn}.xlsx", paths)


def validate_paths(ctx) :
    missing = [str(path) for path in ctx.paths.values() if not path.exists()]
    if missing:
        raise FileNotFoundError("以下月度源文件不存在：\n" + "\n".join(missing))
    log(f"源文件存在性校验完成，共 {len(ctx.paths)} 个文件")


XML_NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def column_number(reference) :
    letters = re.match(r"[A-Z]+", reference).group(0)
    value = 0
    for letter in letters:
        value = value * 26 + ord(letter) - 64
    return value


def scalar(value, cell_type, shared) :
    if value is None:
        return None
    if cell_type == "s":
        return shared[int(value)]
    if cell_type == "b":
        return value == "1"
    if cell_type == "e":
        return value
    if cell_type in ("str", "inlineStr"):
        return value
    try:
        number = float(value)
        return int(number) if number.is_integer() else number
    except ValueError:
        return value


def workbook_sheet_paths(archive) :
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    rel_map = {node.get("Id"): node.get("Target") for node in relationships}
    rel_ns = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
    result = []
    for sheet in workbook.findall(".//m:sheets/m:sheet", XML_NS):
        target = rel_map[sheet.get(rel_ns)].lstrip("/")
        if not target.startswith("xl/"):
            target = "xl/" + target
        result.append((sheet.get("name"), target))
    return result


def read_xlsx_xml(path, sheet_name = None) :
    """Read the first worksheet without trusting its often-broken A1 dimension metadata."""
    with zipfile.ZipFile(path) as archive:
        shared = []
        if "xl/sharedStrings.xml" in archive.namelist():
            root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            for item in root.findall("m:si", XML_NS):
                shared.append("".join(node.text or "" for node in item.iterfind(".//m:t", XML_NS)))
        sheet_paths = workbook_sheet_paths(archive)
        if not sheet_paths:
            raise ValueError(f"工作簿没有工作表 XML：{path}")
        selected = sheet_paths[0][1] if sheet_name is None else dict(sheet_paths).get(sheet_name)
        if selected is None:
            raise KeyError(f"工作簿缺少 Sheet：{sheet_name}")
        root = ET.fromstring(archive.read(selected))
        result = []
        for row_node in root.findall(".//m:sheetData/m:row", XML_NS):
            sparse = {}
            for cell in row_node.findall("m:c", XML_NS):
                col = column_number(cell.get("r"))
                typ = cell.get("t")
                if typ == "inlineStr":
                    value = "".join(node.text or "" for node in cell.iterfind(".//m:t", XML_NS))
                else:
                    node = cell.find("m:v", XML_NS)
                    value = scalar(None if node is None else node.text, typ, shared)
                sparse[col] = value
            if sparse:
                result.append([sparse.get(col) for col in range(1, max(sparse) + 1)])
        return result


def normalized(value) :
    return "" if value is None else str(value).replace("\xa0", " ").strip()


def header_positions(headers) :
    return {normalized(value): index for index, value in enumerate(headers, start=1) if normalized(value)}


def write_value(cell, value) :
    cell.value = value
    if isinstance(value, str) and value.startswith("#"):
        cell.data_type = "e"


def create_workbook() :
    wb = Workbook()
    del wb[wb.sheetnames[0]]
    for name in SHEET_ORDER:
        wb.create_sheet(name)
    return wb


def style_sheet(ws, formula_columns = ()) :
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    ws.row_dimensions[1].height = 28
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = CELL_BORDER
    for index, header in enumerate(ws[1], start=1):
        text = normalized(header.value)
        ws.column_dimensions[get_column_letter(index)].width = min(max(len(text) * 1.7 + 3, 11), 28)


def copy_value_sheet(source_path, source_name, target, wanted_headers = None) :
    wb = load_workbook(source_path, read_only=True, data_only=True)
    source = wb[source_name]
    source_headers = [source.cell(1, col).value for col in range(1, source.max_column + 1)]
    positions = header_positions(source_headers)
    selected = wanted_headers or [normalized(value) for value in source_headers]
    target.append(selected)
    count = 0
    for row in source.iter_rows(min_row=2, values_only=True):
        values = [row[positions[name] - 1] if positions.get(name) and positions[name] <= len(row) else None for name in selected]
        if not any(value not in (None, "") for value in values):
            continue
        target.append(values)
        count += 1
    wb.close()
    return count


def generate_scope_and_mapping(wb, ctx) :
    for target_name, source_name in GRAIN_SOURCES.items():
        rows = copy_value_sheet(ctx.paths["grain"], source_name, wb[target_name], HEADERS[target_name])
        style_sheet(wb[target_name])
        log(f"Sheet完成：{target_name}，数据行={rows}，字段数={len(HEADERS[target_name])}")
    for name in ("SKU信息表", "店铺信息表", "仓库信息表"):
        wanted = None if name == "SKU信息表" else HEADERS[name]
        rows = copy_value_sheet(ctx.paths["mapping"], name, wb[name], wanted)
        style_sheet(wb[name])
        log(f"Sheet完成：{name}，数据行={rows}，字段数={wb[name].max_column}")


def generate_sales(wb, ctx) :
    rows = read_xlsx_xml(ctx.paths["sales"])
    if not rows or normalized(rows[0][0]) != "系统sku":
        raise ValueError("商品销售统计未识别到预期表头")
    ws = wb["千易商品销售统计"]
    ws.append(HEADERS[ws.title])
    for source in rows[1:]:
        if not source or source[0] in (None, ""):
            continue
        source += [None] * (28 - len(source))
        target_row = ws.max_row + 1
        output = [None, None, source[0], source[1], source[2], source[3], source[4], None, source[5], source[6], source[7], None, source[8], None, source[9], source[10], source[11], None, source[12], source[13], None, source[14], source[15], source[16], source[17], source[18], source[19], source[20], None, source[21], source[22], source[23], source[24], source[25], source[26], source[27], None]
        ws.append(output)
        ws.cell(target_row, 37, f'=DATE(LEFT(AJ{target_row},4),MID(AJ{target_row},5,2),1)')
        ws.cell(target_row, 1, f'=TEXT(AK{target_row},"yyyy年mm月")')
        ws.cell(target_row, 2, f'="{ctx.month.year}年第"&ROUNDUP(MONTH(AK{target_row})/3,0)&"季度"')
        ws.cell(target_row, 8, f'=IFERROR(XLOOKUP(C{target_row},SKU信息表!F:F,SKU信息表!L:L),"")')
        ws.cell(target_row, 12, f'=IFERROR(XLOOKUP(K{target_row},店铺信息表!B:B,店铺信息表!C:C),"")')
        ws.cell(target_row, 14, f'=IFERROR(XLOOKUP(M{target_row},仓库信息表!C:C,仓库信息表!B:B),"")')
        ws.cell(target_row, 18, f'=IFERROR(Q{target_row}+0,0)')
        ws.cell(target_row, 21, f'=IFERROR(T{target_row}+0,0)')
        ws.cell(target_row, 29, f'=IFERROR(AB{target_row}+0,0)')
    style_sheet(ws, (1, 2, 8, 12, 14, 18, 21, 29, 37))
    ws.column_dimensions["AJ"].width = 22
    log(f"Sheet完成：{ws.title}，数据行={ws.max_row - 1}，字段数={ws.max_column}，公式列=9")


def generate_receiving(wb, ctx) :
    rows = read_xlsx_xml(ctx.paths["receiving"])
    source_headers = [normalized(value) for value in rows[0]]
    required = {"批注", "预计数量", "仓库上架数量", "商品编码", "收货仓"}
    if not required.issubset(set(source_headers)):
        raise ValueError(f"收货中源表缺少字段：{sorted(required - set(source_headers))}")
    ws = wb["千易入库单_收货中"]
    ws.append(HEADERS[ws.title])
    target_positions = header_positions(HEADERS[ws.title])
    remark_letter = get_column_letter(target_positions["批注"])
    expected_letter = get_column_letter(target_positions["预计数量"])
    shelved_letter = get_column_letter(target_positions["仓库上架数量"])
    difference_col = target_positions["差异数据"]
    for source in rows[1:]:
        source += [None] * (46 - len(source))
        if not any(value not in (None, "") for value in source):
            continue
        row = ws.max_row + 1
        ws.append(source[:32] + [None] + source[32:46])
        ws.cell(
            row,
            difference_col,
            f'=IF(ISNUMBER(FIND("点完",{remark_letter}{row})),MAX({expected_letter}{row}-{shelved_letter}{row},0),0)',
        )
    style_sheet(ws, (difference_col,))
    log(f"Sheet完成：{ws.title}，数据行={ws.max_row - 1}，字段数={ws.max_column}，公式列=差异数据")


def generate_stock(wb, ctx) :
    rows = read_xlsx_xml(ctx.paths["stock"])
    if len(rows) < 3 or normalized(rows[1][0]) != "商品编码":
        raise ValueError("库存查询未识别到第二行字段表头")
    ws = wb["千易库存数据"]
    ws.append(HEADERS[ws.title])
    for source in rows[2:]:
        source += [None] * (11 - len(source))
        if source[0] in (None, ""):
            continue
        row = ws.max_row + 1
        ws.append([ctx.month_cn, ctx.quarter_cn, source[0], source[1], None, source[2], None, source[3], source[4], None, None, None] + source[5:11])
        ws.cell(row, 5, f'=IFERROR(XLOOKUP(C{row},SKU信息表!F:F,SKU信息表!L:L),"")')
        ws.cell(row, 7, f'=IFERROR(XLOOKUP(F{row},仓库信息表!A:A,仓库信息表!B:B),"")')
        ws.cell(row, 10, f'=SUMIFS(\'千易入库单_收货中\'!AG:AG,\'千易入库单_收货中\'!P:P,C{row},\'千易入库单_收货中\'!G:G,F{row})')
        ws.cell(row, 11, f'=MAX(I{row}-J{row},0)')
        ws.cell(row, 12, f'=H{row}+K{row}')
    style_sheet(ws, (5, 7, 10, 11, 12))
    log(f"Sheet完成：{ws.title}，数据行={ws.max_row - 1}，字段数={ws.max_column}，未收到库存粒度=商品编码+仓库")


def demand_header_row(ws) :
    for row in range(1, min(ws.max_row, 6) + 1):
        if normalized(ws.cell(row, 6).value) == "产品编码":
            return row
    raise ValueError(f"{ws.title} 未识别到产品编码表头")


def demand_header_index(rows) :
    for index, row in enumerate(rows[:6]):
        if len(row) >= 6 and normalized(row[5]) == "产品编码":
            return index
    raise ValueError("净需求表未识别到产品编码表头")


def month_target_headers(month) :
    class_targets = [f"类目目标{add_months(month, offset).year}.{add_months(month, offset).month:02d}" for offset in range(1, 7)]
    sku_targets = [f"SKU目标{add_months(month, offset).year}.{add_months(month, offset).month:02d}" for offset in range(1, 7)]
    helper_targets = [f"类目目标{month.year}.{month.month}月", f"SKU目标{month.year}.{month.month}月"]
    return ["站点(CN)", "站点(EN)", "主SKU分类", "材质", "子类目", "菲律宾产品编码", "产品编码", "可售", "实际在途"] + class_targets + sku_targets + ["SKU在途批次计数(月份限制)", "SKU周转系数", "类目在途批次计数", "类目周转系数", "在途数量1", "到货日期1", "日期系数1", "在途数量2", "到货日期2", "日期系数2", "在途数量3", "到货日期3", "日期系数3", "在途数量4", "到货日期4", "日期系数4"] + helper_targets


def demand_helper_headers(month) :
    return f"类目目标{month.year}.{month.month}月", f"SKU目标{month.year}.{month.month}月"


def generate_demand_target(wb, ctx) :
    ws = wb["净需求-目标表"]
    ws.append(month_target_headers(ctx.month))
    imported_errors = Counter()
    helper_class_header, helper_sku_header = demand_helper_headers(ctx.month)
    for site, code in SITE_CODES.items():
        path = ctx.paths[f"demand_{site}"]
        source_rows = read_xlsx_xml(path)
        header_index = demand_header_index(source_rows)
        source_headers = [normalized(value) for value in source_rows[header_index]]
        try:
            helper_class_index = source_headers.index(helper_class_header)
            helper_sku_index = source_headers.index(helper_sku_header)
        except ValueError as exc:
            raise ValueError(f"{path.name} / 净需求表缺少辅助目标字段：{helper_class_header}、{helper_sku_header}") from exc
        count = 0
        for source in source_rows[header_index + 1:]:
            source += [None] * (max(44, helper_sku_index + 1) - len(source))
            product = source[5]
            if product in (None, ""):
                continue
            values = [site.replace("印尼1", "印尼"), code]
            values += source[1:6]
            values += [source[11], source[12]]
            values += source[15:27]
            values += [None, None, None, None]
            values += source[32:44]
            values += [source[helper_class_index], source[helper_sku_index]]
            target_row = ws.max_row + 1
            ws.append([None] * len(values))
            for col, value in enumerate(values, start=1):
                if isinstance(value, str) and value.startswith("#"):
                    imported_errors[value] += 1
                write_value(ws.cell(target_row, col), value)
            aa, ad, ag, aj = (f"AA{target_row}", f"AD{target_row}", f"AG{target_row}", f"AJ{target_row}")
            z, ac, af, ai = (f"Z{target_row}", f"AC{target_row}", f"AF{target_row}", f"AI{target_row}")
            ws.cell(target_row, 22, f'=IFS(OR(A{target_row}="菲律宾",A{target_row}="马来",A{target_row}="泰国"),IF(AND({aa}>=EDATE(\'千易商品销售统计\'!$AK$2,1),{aa}<EDATE(\'千易商品销售统计\'!$AK$2,4)),{z},0)+IF(AND({ad}>=EDATE(\'千易商品销售统计\'!$AK$2,1),{ad}<EDATE(\'千易商品销售统计\'!$AK$2,4)),{ac},0)+IF(AND({ag}>=EDATE(\'千易商品销售统计\'!$AK$2,1),{ag}<EDATE(\'千易商品销售统计\'!$AK$2,4)),{af},0)+IF(AND({aj}>=EDATE(\'千易商品销售统计\'!$AK$2,1),{aj}<EDATE(\'千易商品销售统计\'!$AK$2,4)),{ai},0),A{target_row}="印尼",IF(AND({aa}>=EDATE(\'千易商品销售统计\'!$AK$2,1),{aa}<EDATE(\'千易商品销售统计\'!$AK$2,5)),{z},0)+IF(AND({ad}>=EDATE(\'千易商品销售统计\'!$AK$2,1),{ad}<EDATE(\'千易商品销售统计\'!$AK$2,5)),{ac},0)+IF(AND({ag}>=EDATE(\'千易商品销售统计\'!$AK$2,1),{ag}<EDATE(\'千易商品销售统计\'!$AK$2,5)),{af},0)+IF(AND({aj}>=EDATE(\'千易商品销售统计\'!$AK$2,1),{aj}<EDATE(\'千易商品销售统计\'!$AK$2,5)),{ai},0),TRUE,"")')
            ws.cell(target_row, 23, f'=IFERROR((H{target_row}+I{target_row}+V{target_row})/IFS(OR(A{target_row}="菲律宾",A{target_row}="马来",A{target_row}="泰国"),P{target_row}+Q{target_row}+R{target_row},A{target_row}="印尼",P{target_row}+Q{target_row}+R{target_row}+S{target_row},TRUE,""),"计算错误")')
            ws.cell(target_row, 24, f'=SUMIFS($H:$H,$A:$A,A{target_row},$C:$C,C{target_row})+SUMIFS($I:$I,$A:$A,A{target_row},$C:$C,C{target_row})+SUMIFS($V:$V,$A:$A,A{target_row},$C:$C,C{target_row})')
            ws.cell(target_row, 25, f'=IFERROR(X{target_row}/IFS(OR(A{target_row}="菲律宾",A{target_row}="马来",A{target_row}="泰国"),SUMIFS($P:$P,$A:$A,A{target_row},$C:$C,C{target_row})+SUMIFS($Q:$Q,$A:$A,A{target_row},$C:$C,C{target_row})+SUMIFS($R:$R,$A:$A,A{target_row},$C:$C,C{target_row}),A{target_row}="印尼",SUMIFS($P:$P,$A:$A,A{target_row},$C:$C,C{target_row})+SUMIFS($Q:$Q,$A:$A,A{target_row},$C:$C,C{target_row})+SUMIFS($R:$R,$A:$A,A{target_row},$C:$C,C{target_row})+SUMIFS($S:$S,$A:$A,A{target_row},$C:$C,C{target_row}),TRUE,"计算错误"),"")')
            count += 1
        log(f"净需求导入：{site}，有效产品行={count}")
    style_sheet(ws, (22, 23, 24, 25))
    for col in (27, 30, 33, 36):
        for cell in ws.iter_cols(min_col=col, max_col=col, min_row=2):
            cell[0].number_format = "yyyy-mm-dd"
    log(f"Sheet完成：{ws.title}，数据行={ws.max_row - 1}，忽略印尼2，原样错误={dict(imported_errors)}")


def copy_scope_dimensions(scope, target, dimension_map) :
    source_positions = header_positions([cell.value for cell in scope[1]])
    target.append(HEADERS[target.title])
    for source_row in range(2, scope.max_row + 1):
        values = []
        for target_header in HEADERS[target.title]:
            source_header = dict(dimension_map).get(target_header)
            values.append(scope.cell(source_row, source_positions[source_header]).value if source_header else None)
        target.append(values)


def generate_business_tables(wb, ctx) :
    configs = [
        ("A1类目销售数据表", "A1类目销售数据口径", [("供应商", "供应商"), ("站点", "站点（中文）"), ("产品分级", "产品分级"), ("产品主类/材质", "产品主类/材质"), ("主SKU分类", "主SKU分类")]),
        ("A2类目SKU销售监控表", "A2类目SKU销售监控数据口径", [("供应商", "供应商"), ("站点", "站点（中文）"), ("产品分级", "产品分级"), ("产品主类/材质", "产品主类/材质"), ("主SKU分类", "主SKU分类"), ("产品编码", "产品编码"), ("产品名称", "产品名称")]),
        ("B2店铺类目销售表", "B2店铺类目数据口径", [("站点", "站点"), ("店铺名", "店铺名"), ("运营", "运营"), ("产品分级", "产品分级"), ("产品主类/材质", "产品主类/材质"), ("主SKU分类", "主SKU分类")]),
        ("B1运营类目销售表", "B1运营类目数据口径", [("站点", "站点（中文）"), ("产品分级", "产品分级"), ("产品主类/材质", "产品主类/材质"), ("主SKU分类", "类目"), ("运营", "运营")]),
    ]
    for target_name, scope_name, dimensions in configs:
        ws = wb[target_name]
        copy_scope_dimensions(wb[scope_name], ws, dimensions)
        for row in range(2, ws.max_row + 1):
            if target_name == "A1类目销售数据表":
                formulas = {6: f'="{ctx.month_cn}"', 7: f'="{ctx.quarter_cn}"', 8: f'=SUMIFS(\'千易商品销售统计\'!U:U,\'千易商品销售统计\'!H:H,E{row},\'千易商品销售统计\'!N:N,B{row})', 9: f'=SUMIFS(\'千易商品销售统计\'!R:R,\'千易商品销售统计\'!H:H,E{row},\'千易商品销售统计\'!N:N,B{row})', 10: f'=IFERROR(I{row}/SUMIFS(I:I,B:B,B{row}),0)', 11: f'=XLOOKUP(B{row}&E{row}&"",\'净需求-目标表\'!$A:$A&\'净需求-目标表\'!$C:$C&"",\'净需求-目标表\'!$AL:$AL,"")', 12: f'=SUMIFS(\'千易库存数据\'!L:L,\'千易库存数据\'!G:G,B{row},\'千易库存数据\'!E:E,E{row})', 13: f'=IFERROR(L{row}/K{row},"")', 14: f'=IFERROR(AVERAGEIFS(\'净需求-目标表\'!Y:Y,\'净需求-目标表\'!A:A,B{row},\'净需求-目标表\'!C:C,E{row}),"")'}
            elif target_name == "A2类目SKU销售监控表":
                formulas = {8: f'="{ctx.month_cn}"', 9: f'="{ctx.quarter_cn}"', 10: f'=SUMIFS(\'净需求-目标表\'!AM:AM,\'净需求-目标表\'!A:A,B{row},\'净需求-目标表\'!G:G,F{row})', 11: f'=SUMIFS(\'千易商品销售统计\'!U:U,\'千易商品销售统计\'!C:C,F{row},\'千易商品销售统计\'!N:N,B{row})', 12: f'=IFERROR(K{row}/J{row},"")', 13: f'=SUMIFS(\'千易库存数据\'!L:L,\'千易库存数据\'!G:G,B{row},\'千易库存数据\'!C:C,F{row})', 14: f'=IFERROR(M{row}/J{row},"")', 15: f'=IFERROR(AVERAGEIFS(\'净需求-目标表\'!W:W,\'净需求-目标表\'!A:A,B{row},\'净需求-目标表\'!G:G,F{row}),"")'}
            elif target_name == "B2店铺类目销售表":
                formulas = {7: f'="{ctx.month_cn}"', 8: f'="{ctx.quarter_cn}"', 9: f'=SUMIFS(\'千易库存数据\'!L:L,\'千易库存数据\'!G:G,A{row},\'千易库存数据\'!E:E,F{row})', 10: f'=SUMIFS(\'千易商品销售统计\'!U:U,\'千易商品销售统计\'!N:N,A{row},\'千易商品销售统计\'!K:K,B{row},\'千易商品销售统计\'!H:H,F{row})', 11: f'=IFERROR(AVERAGEIFS(\'千易商品销售统计\'!AC:AC,\'千易商品销售统计\'!N:N,A{row},\'千易商品销售统计\'!K:K,B{row},\'千易商品销售统计\'!H:H,F{row},\'千易商品销售统计\'!AC:AC,">0"),0)', 12: f'=SUMIFS(\'千易商品销售统计\'!R:R,\'千易商品销售统计\'!N:N,A{row},\'千易商品销售统计\'!K:K,B{row},\'千易商品销售统计\'!H:H,F{row})'}
            else:
                formulas = {6: f'="{ctx.month_cn}"', 7: f'="{ctx.quarter_cn}"', 8: f'=SUMIFS(\'千易商品销售统计\'!U:U,\'千易商品销售统计\'!N:N,A{row},\'千易商品销售统计\'!L:L,E{row},\'千易商品销售统计\'!H:H,D{row})', 9: f'=SUMIFS(\'千易商品销售统计\'!R:R,\'千易商品销售统计\'!N:N,A{row},\'千易商品销售统计\'!L:L,E{row},\'千易商品销售统计\'!H:H,D{row})', 10: f'=IFERROR(I{row}/SUMIFS(I:I,A:A,A{row},E:E,E{row}),0)'}
            for col, formula in formulas.items():
                ws.cell(row, col, formula)
        formula_start = len(dimensions) + 1
        style_sheet(ws, range(formula_start, ws.max_column + 1))
        log(f"Sheet完成：{target_name}，数据行={ws.max_row - 1}，字段数={ws.max_column}")


def recalculate_with_excel(path, timeout_seconds = 1200) :
    script = r'''
param([string]$WorkbookPath,[int]$TimeoutSeconds)
$ErrorActionPreference='Stop'; $excel=$null; $book=$null
try {
  $excel=New-Object -ComObject Excel.Application
  $excel.Visible=$false; $excel.DisplayAlerts=$false; $excel.AskToUpdateLinks=$false
  $book=$excel.Workbooks.Open($WorkbookPath,0,$false)
  $formulaColor=16247773
  foreach($sheet in $book.Worksheets){
    try {$formulaCells=$sheet.UsedRange.SpecialCells(-4123); $formulaCells.Interior.Color=$formulaColor}
    catch {}
  }
  $excel.CalculateFullRebuild(); $start=Get-Date
  while($excel.CalculationState -ne 0){if(((Get-Date)-$start).TotalSeconds -gt $TimeoutSeconds){throw 'Excel公式计算超时'}; Start-Sleep -Milliseconds 500}
  $book.Save()
} finally {
  if($book-ne$null){$book.Close($true)|Out-Null;[Runtime.InteropServices.Marshal]::ReleaseComObject($book)|Out-Null}
  if($excel-ne$null){$excel.Quit();[Runtime.InteropServices.Marshal]::ReleaseComObject($excel)|Out-Null}
  [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
'''
    temporary = None
    try:
        with tempfile.NamedTemporaryFile("w", suffix=".ps1", delete=False, encoding="utf-8-sig") as handle:
            handle.write(script)
            temporary = Path(handle.name)
        result = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(temporary), "-WorkbookPath", str(path.resolve()), "-TimeoutSeconds", str(timeout_seconds)], capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=timeout_seconds + 60)
        if result.returncode:
            raise RuntimeError("Excel公式计算失败：" + (result.stderr or result.stdout))
    finally:
        if temporary and temporary.exists():
            temporary.unlink()


def validate_output(ctx) :
    with zipfile.ZipFile(ctx.output) as archive:
        sheet_paths = workbook_sheet_paths(archive)
        sheet_names = [name for name, _ in sheet_paths]
        if sheet_names != SHEET_ORDER:
            raise ValueError(f"Sheet顺序不正确：{sheet_names}")
        ref_formula_files = []
        for name, xml_path in sheet_paths:
            raw = archive.read(xml_path)
            if re.search(br"<f(?:\s[^>]*)?>[^<]*#REF!", raw):
                ref_formula_files.append(name)
        if ref_formula_files:
            raise ValueError("新生成公式包含 #REF!：" + ", ".join(ref_formula_files))
    for target, source in (("A1类目销售数据表", "A1类目销售数据口径"), ("A2类目SKU销售监控表", "A2类目SKU销售监控数据口径"), ("B1运营类目销售表", "B1运营类目数据口径"), ("B2店铺类目销售表", "B2店铺类目数据口径")):
        if len(read_xlsx_xml(ctx.output, target)) != len(read_xlsx_xml(ctx.output, source)):
            raise ValueError(f"{target} 行数与口径表不一致")
    demand_rows = read_xlsx_xml(ctx.output, "净需求-目标表")
    if any(row and row[0] == "印尼2" for row in demand_rows[1:]):
        raise ValueError("净需求-目标表中不应包含印尼2")
    imported_na = sum(1 for row in demand_rows for value in row if value == "#N/A")
    log(f"最终校验通过：15个Sheet，导入并原样保留的#N/A={imported_na}，新公式#REF!=0")


DASHBOARD_CONFIGS = {
    "A1类目销售数据表": {
        "dims": ["供应商", "站点", "产品分级", "产品主类/材质", "主SKU分类", "平均销售额占比"],
        "keys": ["站点", "主SKU分类"],
        "block": ["月份", "季度", "目标", "销量", "目标完成率", "销售额", "目标增长系数", "销售额增长系数", "销售额占比", "计划目标", "实际库存", "备货准确率", "周转系数"],
        "manual": "目标",
    },
    "A2类目SKU销售监控表": {
        "dims": ["供应商", "站点", "产品分级", "产品主类/材质", "主SKU分类", "产品编码", "产品名称"],
        "keys": ["站点", "产品编码"],
        "block": ["月份", "季度", "计划目标", "实际销量", "目标完成率", "实际库存", "备货准确率", "周转系数"],
    },
    "B1运营类目销售表": {
        "dims": ["站点", "产品分级", "产品主类/材质", "主SKU分类", "运营"],
        "keys": ["站点", "运营", "主SKU分类"],
        "block": ["月份", "季度", "实际销量", "目标销量", "销量完成率", "实际销售额", "目标销售额", "销售额完成率", "目标增长系数", "销售额增长系数", "销售额占比"],
    },
    "B2店铺类目销售表": {
        "dims": ["站点", "店铺名", "运营", "产品分级", "产品主类/材质", "主SKU分类"],
        "keys": ["站点", "店铺名", "主SKU分类"],
        "block": ["月份", "季度", "实际库存", "客单价（RMB）", "实际销量", "目标销量（手填）", "销量完成率", "实际销售额", "目标销售额", "销售额完成率", "销量增长系数", "销售额增长系数"],
        "manual": "目标销量（手填）",
    },
}


def dashboard_key(record, key_fields) :
    return tuple(normalized(record.get(field)) for field in key_fields)


def safe_history_value(value) :
    return None if isinstance(value, str) and value.startswith("#") else value


def copy_cell_appearance(source, target) :
    if source.has_style:
        target._style = copy(source._style)
    if source.number_format:
        target.number_format = source.number_format
    target.alignment = copy(source.alignment)
    target.protection = copy(source.protection)


def month_from_value(value) :
    if isinstance(value, (datetime, date)):
        return value.date().replace(day=1) if isinstance(value, datetime) else value.replace(day=1)
    text = normalized(value)
    match = re.search(r"(\d{4})年(\d{1,2})月", text)
    if match:
        return date(int(match.group(1)), int(match.group(2)), 1)
    if isinstance(value, (int, float)):
        converted = from_excel(value)
        return converted.date().replace(day=1) if isinstance(converted, datetime) else converted.replace(day=1)
    return None


def sheet_records(ws, headers) :
    positions = header_positions([cell.value for cell in ws[1]])
    missing = [header for header in headers if header not in positions]
    if missing:
        raise ValueError(f"{ws.title} 缺少字段：{missing}")
    records = []
    for row in range(2, ws.max_row + 1):
        record = {header: ws.cell(row, positions[header]).value for header in headers}
        if any(value not in (None, "") for value in record.values()):
            records.append(record)
    return records


def product_grade_formula(sheet_name, row) :
    if sheet_name in ("A1类目销售数据表", "A2类目SKU销售监控表"):
        site, category = f"B{row}", f"E{row}"
    elif sheet_name == "B1运营类目销售表":
        site, category = f"A{row}", f"D{row}"
    else:
        site, category = f"A{row}", f"F{row}"
    return f'=IFERROR(XLOOKUP({site}&{category},\'产品动态分级表\'!A:A&\'产品动态分级表\'!C:C,\'产品动态分级表\'!Y:Y,"未定级"),"未定级")'


def snapshot_dashboard_sheet(formula_ws, value_ws, config) :
    dims = config["dims"]
    block = config["block"]
    dim_count, block_width = len(dims), len(block)
    total_blocks = max((formula_ws.max_column - dim_count) // block_width, 1)
    positions = {header: index + 1 for index, header in enumerate(dims)}
    snapshots = {}
    order = []
    for row in range(2, formula_ws.max_row + 1):
        record = {header: value_ws.cell(row, positions[header]).value for header in dims}
        key = dashboard_key(record, config["keys"])
        if not any(key) or key in snapshots:
            continue
        blocks = []
        for block_index in range(total_blocks):
            start = dim_count + 1 + block_index * block_width
            blocks.append([safe_history_value(value_ws.cell(row, start + offset).value) for offset in range(block_width)])
        snapshots[key] = {
            "dims": record,
            "blocks": blocks,
            "styles": [copy(formula_ws.cell(row, col)._style) for col in range(1, formula_ws.max_column + 1)],
            "height": formula_ws.row_dimensions[row].height,
        }
        order.append(key)
    return snapshots, order, total_blocks


def rebuild_dashboard_sheet(result_wb, result_values, middle_values, name, ctx) :
    config = DASHBOARD_CONFIGS[name]
    ws = result_wb[name]
    values_ws = result_values[name]
    middle_ws = middle_values[name]
    dims, block = config["dims"], config["block"]
    dim_count, block_width = len(dims), len(block)
    snapshots, old_order, old_block_count = snapshot_dashboard_sheet(ws, values_ws, config)
    old_current_month = None
    for snapshot in snapshots.values():
        old_current_month = month_from_value(snapshot["blocks"][0][0])
        if old_current_month:
            break
    roll_month = old_current_month != ctx.month
    history_count = old_block_count if roll_month else max(old_block_count - 1, 0)
    source_block_months = []
    for block_index in range(old_block_count):
        block_month = None
        for snapshot in snapshots.values():
            block_month = month_from_value(snapshot["blocks"][block_index][0])
            if block_month:
                break
        source_block_months.append(block_month)

    middle_headers = [cell.value for cell in middle_ws[1]]
    middle_positions = header_positions(middle_headers)
    grain_records = []
    for row in range(2, middle_ws.max_row + 1):
        record = {header: middle_ws.cell(row, col).value for header, col in middle_positions.items()}
        if any(value not in (None, "") for value in record.values()):
            grain_records.append(record)
    grain_keys = [dashboard_key(record, config["keys"]) for record in grain_records]
    if len(grain_keys) != len(set(grain_keys)):
        raise ValueError(f"{name} 月度中间表存在重复粒度")
    grain_map = dict(zip(grain_keys, grain_records))
    legacy_keys = [key for key in old_order if key not in grain_map]
    final_keys = grain_keys + legacy_keys

    old_header_styles = [copy(ws.cell(1, col)._style) for col in range(1, ws.max_column + 1)]
    old_widths = [ws.column_dimensions[get_column_letter(col)].width for col in range(1, ws.max_column + 1)]
    template_styles = [copy(ws.cell(2, min(col, ws.max_column))._style) for col in range(1, dim_count + block_width * (1 + history_count) + 1)]
    template_height = ws.row_dimensions[2].height
    ws.delete_rows(1, ws.max_row)
    total_columns = dim_count + block_width * (1 + history_count)
    headers = dims + block * (1 + history_count)
    ws.append(headers)
    for col in range(1, total_columns + 1):
        style_source = col if col <= len(old_header_styles) else dim_count + 1 + ((col - dim_count - 1) % block_width)
        ws.cell(1, col)._style = copy(old_header_styles[style_source - 1])
        width_source = col if col <= len(old_widths) else dim_count + 1 + ((col - dim_count - 1) % block_width)
        ws.column_dimensions[get_column_letter(col)].width = old_widths[width_source - 1]

    for output_row, key in enumerate(final_keys, start=2):
        grain = grain_map.get(key)
        snapshot = snapshots.get(key)
        legacy = grain is None
        if grain:
            for col, header in enumerate(dims, start=1):
                if header == "产品分级":
                    ws.cell(output_row, col, product_grade_formula(name, output_row))
                elif header == "平均销售额占比":
                    pass
                else:
                    ws.cell(output_row, col, grain.get(header))
        elif snapshot:
            for col, header in enumerate(dims, start=1):
                ws.cell(output_row, col, snapshot["dims"].get(header))
                if header == "产品分级":
                    ws.cell(output_row, col, product_grade_formula(name, output_row))

        current_start = dim_count + 1
        ws.cell(output_row, current_start, ctx.month_cn)
        ws.cell(output_row, current_start + 1, ctx.quarter_cn)
        manual_header = config.get("manual")
        manual_value = None
        if manual_header and snapshot:
            manual_value = snapshot["blocks"][0][block.index(manual_header)]
        if manual_header:
            ws.cell(output_row, current_start + block.index(manual_header), manual_value)

        if not legacy:
            if name == "A1类目销售数据表":
                mapping = {"销量": "销量", "销售额": "销售额", "计划目标": "计划目标", "实际库存": "实际库存", "备货准确率": "备货准确率", "周转系数": "周转系数"}
            elif name == "A2类目SKU销售监控表":
                mapping = {"计划目标": "计划目标", "实际销量": "实际销量", "实际库存": "实际库存", "周转系数": "周转系数"}
            elif name == "B2店铺类目销售表":
                mapping = {"实际库存": "实际库存", "客单价（RMB）": "客单价（RMB）", "实际销量": "实际销量", "实际销售额": "实际销售额"}
            else:
                mapping = {"实际销量": "实际销量", "实际销售额": "实际销售额"}
            for target_header, source_header in mapping.items():
                ws.cell(output_row, current_start + block.index(target_header), grain.get(source_header))

            if name == "A1类目销售数据表":
                ws.cell(output_row, 6, f'=IFERROR(AVERAGE(O{output_row}' + ''.join(f',{get_column_letter(dim_count + 1 + index * block_width + block.index("销售额占比"))}{output_row}' for index in range(1, 1 + history_count)) + '),0)')
                ws.cell(output_row, current_start + 4, f'=IFERROR(J{output_row}/I{output_row},"")')
                history_start = current_start + block_width
                ws.cell(output_row, current_start + 6, f'=IFERROR(I{output_row}/{get_column_letter(history_start + 2)}{output_row},"")' if history_count else None)
                ws.cell(output_row, current_start + 7, f'=IFERROR(L{output_row}/{get_column_letter(history_start + 5)}{output_row},"")' if history_count else None)
                ws.cell(output_row, current_start + 8, f'=IFERROR(L{output_row}/SUMIFS(L:L,B:B,B{output_row}),0)')
            elif name == "A2类目SKU销售监控表":
                ws.cell(output_row, current_start + 4, f'=IFERROR(K{output_row}/J{output_row},"")')
                ws.cell(output_row, current_start + 6, f'=IFERROR(M{output_row}/J{output_row},"")')
            elif name == "B2店铺类目销售表":
                ws.cell(output_row, current_start + 6, f'=IFERROR(K{output_row}/L{output_row},"")')
                ws.cell(output_row, current_start + 8, f'=IFERROR(L{output_row}*J{output_row},0)')
                ws.cell(output_row, current_start + 9, f'=IFERROR(N{output_row}/O{output_row},"")')
                history_start = current_start + block_width
                ws.cell(output_row, current_start + 10, f'=IFERROR(K{output_row}/{get_column_letter(history_start + 4)}{output_row},"")' if history_count else None)
                ws.cell(output_row, current_start + 11, f'=IFERROR(N{output_row}/{get_column_letter(history_start + 7)}{output_row},"")' if history_count else None)
            else:
                ws.cell(output_row, current_start + 3, f'=SUMIFS(\'B2店铺类目销售表\'!L:L,\'B2店铺类目销售表\'!A:A,A{output_row},\'B2店铺类目销售表\'!F:F,D{output_row},\'B2店铺类目销售表\'!C:C,E{output_row})')
                ws.cell(output_row, current_start + 4, f'=IFERROR(H{output_row}/I{output_row},"")')
                ws.cell(output_row, current_start + 6, f'=SUMIFS(\'B2店铺类目销售表\'!O:O,\'B2店铺类目销售表\'!A:A,A{output_row},\'B2店铺类目销售表\'!F:F,D{output_row},\'B2店铺类目销售表\'!C:C,E{output_row})')
                ws.cell(output_row, current_start + 7, f'=IFERROR(K{output_row}/L{output_row},"")')
                history_start = current_start + block_width
                ws.cell(output_row, current_start + 8, f'=IFERROR(I{output_row}/{get_column_letter(history_start + 3)}{output_row},"")' if history_count else None)
                ws.cell(output_row, current_start + 9, f'=IFERROR(L{output_row}/{get_column_letter(history_start + 6)}{output_row},"")' if history_count else None)
                ws.cell(output_row, current_start + 10, f'=IFERROR(K{output_row}/SUMIFS(K:K,A:A,A{output_row},E:E,E{output_row}),0)')

        history_blocks = []
        if snapshot:
            if roll_month:
                history_blocks = snapshot["blocks"]
            else:
                history_blocks = snapshot["blocks"][1:]
        for block_index in range(history_count):
            values = history_blocks[block_index] if block_index < len(history_blocks) else [None] * block_width
            source_block_index = block_index if roll_month else block_index + 1
            block_month = source_block_months[source_block_index] if source_block_index < len(source_block_months) else None
            if block_month:
                values = list(values)
                values[0] = f"{block_month.year}年{block_month.month:02d}月"
                values[1] = f"{block_month.year}年第{(block_month.month - 1) // 3 + 1}季度"
            start = current_start + block_width * (block_index + 1)
            for offset, value in enumerate(values):
                write_value(ws.cell(output_row, start + offset), safe_history_value(value))

        styles = snapshot["styles"] if snapshot else template_styles
        for col in range(1, total_columns + 1):
            style_col = col if col <= len(styles) else dim_count + 1 + ((col - dim_count - 1) % block_width)
            ws.cell(output_row, col)._style = copy(styles[style_col - 1])
        ws.row_dimensions[output_row].height = snapshot["height"] if snapshot else template_height
    ws.auto_filter.ref = ws.dimensions
    log(f"结果表Sheet完成：{name}，有效粒度={len(grain_keys)}，追加旧粒度={len(legacy_keys)}，历史块={history_count}")


def refresh_product_grades(wb) :
    ws = wb["产品动态分级表"]
    for row in range(3, ws.max_row + 1):
        for col in range(5, 26):
            value = ws.cell(row, col).value
            if isinstance(value, str) and value.startswith("#"):
                ws.cell(row, col).value = None
            elif isinstance(value, str) and value.startswith("=") and not value.upper().startswith("=IFERROR("):
                ws.cell(row, col, f'=IFERROR({value[1:]},"")')
    log(f"结果表Sheet完成：产品动态分级表，保留现有粒度={max(ws.max_row - 2, 0)}，外部公式不恢复")


def apply_dashboard_percentage_formats(wb) :
    business_keywords = ("完成率", "增长率", "增长系数", "占比", "备货准确率")
    product_keywords = ("占比", "毛利率", "净利率", "周转率", "权重")
    formatted_columns = 0
    for ws in wb.worksheets:
        header_row = 2 if ws.title == "产品动态分级表" else 1
        first_data_row = header_row + 1
        for col in range(1, ws.max_column + 1):
            header = normalized(ws.cell(header_row, col).value)
            if ws.title == "产品动态分级表":
                is_percentage = any(keyword in header for keyword in product_keywords) and "评分" not in header and "加权分" not in header
            else:
                is_percentage = any(keyword in header for keyword in business_keywords)
            if not is_percentage:
                continue
            for row in range(first_data_row, ws.max_row + 1):
                ws.cell(row, col).number_format = "0.00%"
            formatted_columns += 1
    log(f"结果表百分比格式完成：字段列={formatted_columns}，格式=0.00%")


def update_dashboard(ctx) :
    result_path = ctx.paths["dashboard"]
    temporary_result = result_path.with_name("运营大盘数据监控表_更新中.xlsx")
    backup_dir = result_path.parent / "备份"
    backup_dir.mkdir(parents=True, exist_ok=True)
    backup = backup_dir / f"运营大盘数据监控表_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    shutil.copy2(result_path, backup)
    log(f"结果表已备份：{backup}")
    result_wb = load_workbook(result_path, data_only=False, read_only=False)
    result_values = load_workbook(result_path, data_only=True, read_only=False)
    middle_values = load_workbook(ctx.output, data_only=True, read_only=False)
    try:
        refresh_product_grades(result_wb)
        for name in ("A1类目销售数据表", "A2类目SKU销售监控表", "B2店铺类目销售表", "B1运营类目销售表"):
            rebuild_dashboard_sheet(result_wb, result_values, middle_values, name, ctx)
        apply_dashboard_percentage_formats(result_wb)
        result_wb.calculation.fullCalcOnLoad = True
        result_wb.calculation.forceFullCalc = True
        result_wb.save(temporary_result)
    finally:
        result_wb.close()
        result_values.close()
        middle_values.close()
    recalculate_with_excel(temporary_result)
    validate_dashboard(ctx, temporary_result)
    try:
        temporary_result.replace(result_path)
    except PermissionError as exc:
        raise PermissionError(f"结果表正被WPS或其他程序占用，已生成并校验临时结果：{temporary_result}。请关闭占用后重新执行。") from exc
    log("运营大盘数据监控表已完成Excel全量计算并替换正式文件")
    return backup


def validate_dashboard(ctx, result_path = None) :
    result_path = result_path or ctx.paths["dashboard"]
    with zipfile.ZipFile(result_path) as archive:
        names = [name for name, _ in workbook_sheet_paths(archive)]
        expected = ["产品动态分级表", "A1类目销售数据表", "A2类目SKU销售监控表", "B1运营类目销售表", "B2店铺类目销售表"]
        missing = [name for name in expected if name not in names]
        if missing:
            raise ValueError(f"结果表缺少必需Sheet：{missing}；实际Sheet：{names}")
        bad = []
        for name, xml_path in workbook_sheet_paths(archive):
            if name not in expected:
                continue
            raw = archive.read(xml_path)
            if re.search(br"<f(?:\s[^>]*)?>[^<]*#REF!", raw):
                bad.append(name)
            root = ET.fromstring(raw)
            unexpected_errors = []
            for cell in root.findall(".//m:c[@t='e']", XML_NS):
                value = cell.find("m:v", XML_NS)
                error = None if value is None else value.text
                if error != "#N/A":
                    unexpected_errors.append(f"{cell.get('r')}={error}")
            if unexpected_errors:
                raise ValueError(f"结果表{name}存在非预期缓存错误：{unexpected_errors[:20]}")
        if bad:
            raise ValueError("结果表新公式包含#REF!：" + ",".join(bad))
    for name, config in DASHBOARD_CONFIGS.items():
        rows = read_xlsx_xml(result_path, name)
        middle_rows = read_xlsx_xml(ctx.output, name)
        headers = rows[0]
        month_col = len(config["dims"])
        months = {row[month_col] for row in rows[1:] if len(row) > month_col and row[month_col] not in (None, "")}
        if months != {ctx.month_cn}:
            raise ValueError(f"{name} 当前月份异常：{months}")
        positions = header_positions(headers[:len(config["dims"])])
        keys = [dashboard_key({field: row[positions[field] - 1] for field in config["keys"]}, config["keys"]) for row in rows[1:]]
        if len(keys) != len(set(keys)):
            raise ValueError(f"{name} 存在重复粒度")
        middle_positions = header_positions(middle_rows[0])
        middle_keys = [dashboard_key({field: row[middle_positions[field] - 1] for field in config["keys"]}, config["keys"]) for row in middle_rows[1:]]
        if keys[:len(middle_keys)] != middle_keys:
            raise ValueError(f"{name} 有效粒度顺序未与月度中间表一致")
        dim_count, block_width = len(config["dims"]), len(config["block"])
        block_count = (len(headers) - dim_count) // block_width
        previous = None
        for block_index in range(block_count):
            col = dim_count + block_index * block_width
            block_months = {month_from_value(row[col]) for row in rows[1:] if len(row) > col and row[col] not in (None, "")}
            block_months.discard(None)
            if len(block_months) != 1:
                raise ValueError(f"{name} 第{block_index + 1}个月份块不唯一：{block_months}")
            current = next(iter(block_months))
            if previous and current >= previous:
                raise ValueError(f"{name} 历史月份未按从晚到早排列")
            previous = current
    log("运营大盘最终校验通过：核心Sheet的当前月份、粒度唯一性、结构和公式错误均正常")


def run(root, month) :
    ctx = build_context(root, month)
    log(f"执行参数：ROOT={ctx.root}，M={month}，输出={ctx.output}")
    try:
        validate_paths(ctx)
        wb = create_workbook()
        try:
            generate_scope_and_mapping(wb, ctx)
            generate_sales(wb, ctx)
            generate_receiving(wb, ctx)
            generate_stock(wb, ctx)
            generate_demand_target(wb, ctx)
            generate_business_tables(wb, ctx)
            wb.calculation.fullCalcOnLoad = True
            wb.calculation.forceFullCalc = True
            ctx.output.parent.mkdir(parents=True, exist_ok=True)
            wb.save(ctx.output)
        finally:
            wb.close()
        log("工作簿已写入，开始调用Microsoft Excel全量计算")
        started = time.time()
        recalculate_with_excel(ctx.output)
        log(f"Excel公式计算并保存完成，耗时={time.time() - started:.1f}秒")
        validate_output(ctx)
        update_dashboard(ctx)
        validate_dashboard(ctx)
    finally:
        release_resources_and_verify([ctx.output, ctx.paths["dashboard"]])
    return ctx.output


def parse_args() :
    parser = argparse.ArgumentParser(description="生成月度数据中间表")
    parser.add_argument("--root", help="项目根目录")
    parser.add_argument("--month", help="执行月份，格式YYYY-MM")
    return parser.parse_args()


VSCODE_ROOT = r"<ROOT>"
VSCODE_MONTH = "2026-06"


def main(root = None, month = None) :
    args = parse_args() if root is None and month is None else argparse.Namespace(root=root, month=month)
    run_root = args.root or VSCODE_ROOT
    run_month = args.month or VSCODE_MONTH
    output = run(run_root, run_month)
    print(f"生成完成：{output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
