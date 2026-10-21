# -*- coding: utf-8 -*-
"""将已核定的表格逐行写成带深海蓝格式的 Excel 工作簿。"""

import argparse
import csv
import io
import json
import os
import re
import uuid
import zipfile
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from xml.sax.saxutils import escape


MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
TYPES = {"文本": 49, "金额": 4, "数值": 164, "日期": 165}
BAD_XML = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\ud800-\udfff]")
BAD_SHEET = re.compile(r"[\\/*?:\[\]]")
TOTAL_WORDS = ("合计", "小计", "总计", "sum", "total", "小計", "合計")


def _xml(value):
    value = str(value)
    if BAD_XML.search(value):
        raise ValueError("文本含有 Excel XML 不允许的字符")
    return escape(value, {'"': "&quot;", "'": "&apos;"})


def _column_letter(number):
    result = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _styles():
    fonts = (
        '<font><sz val="11"/><color theme="1"/><name val="Calibri"/></font>',
        '<font><b/><sz val="11"/><color rgb="00FFFFFF"/><name val="Arial"/></font>',
        '<font><sz val="11"/><color rgb="00000000"/><name val="Arial"/></font>',
        '<font><sz val="11"/><color rgb="000000FF"/><name val="Arial"/></font>',
        '<font><b/><sz val="11"/><color rgb="00000000"/><name val="Arial"/></font>',
    )
    borders = (
        '<border><left/><right/><top/><bottom/><diagonal/></border>',
        '<border><left/><right/><top style="medium"><color rgb="00000000"/></top>'
        '<bottom style="dashed"><color rgb="00808080"/></bottom><diagonal/></border>',
        '<border><left/><right/><top style="dashed"><color rgb="00808080"/></top>'
        '<bottom style="dashed"><color rgb="00808080"/></bottom><diagonal/></border>',
        '<border><left/><right/><top style="dashed"><color rgb="00808080"/></top>'
        '<bottom style="medium"><color rgb="00000000"/></bottom><diagonal/></border>',
    )
    fills = (
        '<fill><patternFill patternType="none"/></fill>',
        '<fill><patternFill patternType="gray125"/></fill>',
        '<fill><patternFill patternType="solid"><fgColor rgb="001F4E79"/></patternFill></fill>',
        '<fill><patternFill patternType="solid"><fgColor rgb="00D6E4F0"/></patternFill></fill>',
    )
    xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>']
    xfs.append('<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyAlignment="1">'
               '<alignment horizontal="center" vertical="center"/></xf>')
    lookup = {}
    for role in ("normal", "last", "total"):
        for kind, fmt in TYPES.items():
            font = 4 if role == "total" else (3 if kind in ("金额", "数值") else 2)
            fill = 3 if role == "total" else 0
            border = 3 if role != "normal" else 2
            align = "right" if kind in ("金额", "数值") else "left"
            lookup[role, kind] = len(xfs)
            xfs.append(f'<xf numFmtId="{fmt}" fontId="{font}" fillId="{fill}" '
                       f'borderId="{border}" xfId="0" applyAlignment="1">'
                       f'<alignment horizontal="{align}" vertical="center"/></xf>')
    styles = (
        XML_HEAD + f'<styleSheet xmlns="{MAIN_NS}">'
        '<numFmts count="2"><numFmt numFmtId="164" formatCode="0.##########"/>'
        '<numFmt numFmtId="165" formatCode="yyyy/mm/dd"/></numFmts>'
        f'<fonts count="{len(fonts)}">{"".join(fonts)}</fonts>'
        f'<fills count="{len(fills)}">{"".join(fills)}</fills>'
        f'<borders count="{len(borders)}">{"".join(borders)}</borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        f'<cellXfs count="{len(xfs)}">{"".join(xfs)}</cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        '</styleSheet>'
    )
    return styles, lookup


def _validate_sheets(sheets):
    if not sheets:
        raise ValueError("至少需要一张工作表")
    seen = set()
    for sheet in sheets:
        name = sheet["名称"]
        if not isinstance(name, str) or not name or len(name) > 31 or BAD_SHEET.search(name):
            raise ValueError(f"工作表名称无效：{name!r}")
        if name.casefold() in seen:
            raise ValueError(f"工作表名称重复：{name}")
        seen.add(name.casefold())
        columns = sheet["列"]
        if not columns or len(columns) > 16383:
            raise ValueError(f"{name} 的列数超出范围")
        for column in columns:
            if column["类型"] not in TYPES:
                raise ValueError(f"{name} 存在未知列类型：{column['类型']}")
            _xml(column["名称"])
            width = column.get("宽度", 18 if column["类型"] == "文本" else 16)
            if not isinstance(width, (int, float)) or not 1 <= width <= 255:
                raise ValueError(f"{name} 存在无效列宽：{width}")


def _number(value):
    try:
        number = value if isinstance(value, Decimal) else Decimal(str(value).strip())
    except (InvalidOperation, ValueError) as error:
        raise ValueError(f"数值无效：{value!r}") from error
    if not number.is_finite():
        raise ValueError(f"数值无效：{value!r}")
    if len(number.normalize().as_tuple().digits) > 15:
        raise ValueError(f"数值超过 Excel 可保真的 15 位有效数字：{value!r}")
    if number.adjusted() > 307 or number.adjusted() < -307:
        raise ValueError(f"数值超出 Excel 范围：{value!r}")
    return format(number, "f")


def _date_serial(value):
    if isinstance(value, datetime):
        if value.time().isoformat() != "00:00:00":
            raise ValueError(f"日期包含时间，不能无声截断：{value!r}")
        value = value.date()
    elif not isinstance(value, date):
        try:
            value = date.fromisoformat(str(value))
        except ValueError as error:
            raise ValueError(f"日期必须为 YYYY-MM-DD：{value!r}") from error
    if value < date(1900, 3, 1):
        raise ValueError(f"日期早于 Excel 可稳定表示的范围：{value!r}")
    return str((value - date(1899, 12, 30)).days)


def _cell(ref, value, kind, style):
    prefix = f'<c r="{ref}" s="{style}"'
    if value is None or value == "":
        return prefix + '/>'
    if kind == "文本":
        return prefix + ' t="inlineStr"><is><t xml:space="preserve">' + _xml(value) + '</t></is></c>'
    if kind == "日期":
        value = _date_serial(value)
    else:
        value = _number(value)
    return prefix + f'><v>{value}</v></c>'


def _write_sheet(archive, index, sheet, lookup):
    columns = sheet["列"]
    ncols = len(columns)
    with archive.open(f"xl/worksheets/sheet{index}.xml", "w", force_zip64=True) as output:
        def put(part):
            output.write(part.encode("utf-8"))

        put(XML_HEAD + f'<worksheet xmlns="{MAIN_NS}">')
        put('<sheetViews><sheetView workbookViewId="0" showGridLines="0">'
            '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
            '</sheetView></sheetViews><sheetFormatPr defaultRowHeight="18"/>')
        put('<cols><col min="1" max="1" width="2" customWidth="1"/>')
        for index_col, column in enumerate(columns, 2):
            width = column.get("宽度", 18 if column["类型"] == "文本" else 16)
            put(f'<col min="{index_col}" max="{index_col}" width="{width}" customWidth="1"/>')
        put('</cols><sheetData><row r="1">')
        for index_col, column in enumerate(columns, 2):
            put(_cell(f'{_column_letter(index_col)}1', column["名称"], "文本", 1))
        put('</row>')

        def data_row(values, row_number, last):
            if len(values) != ncols:
                raise ValueError(f'{sheet["名称"]} 第 {row_number} 行列数不符')
            if row_number > 1048576:
                raise ValueError(f'{sheet["名称"]} 超过 Excel 的 1,048,576 行上限')
            label = str(values[0] or "").lower()
            role = "total" if any(word in label for word in TOTAL_WORDS) else ("last" if last else "normal")
            cells = (_cell(f'{_column_letter(col)}{row_number}', value, spec["类型"],
                           lookup[role, spec["类型"]])
                     for col, (value, spec) in enumerate(zip(values, columns), 2))
            put(f'<row r="{row_number}">{"".join(cells)}</row>')

        pending = None
        count = 0
        for values in sheet["行"]:
            if pending is not None:
                data_row(pending, count + 1, False)
            pending = list(values)
            count += 1
        if pending is not None:
            data_row(pending, count + 1, True)
        put('</sheetData><pageMargins left="0.75" right="0.75" top="0.75" bottom="0.75" '
            'header="0.3" footer="0.3"/></worksheet>')
    return count


def write_workbook(sheets, target):
    """逐行写入已整理数据；返回各表数据行数，不计算或改变业务结果。"""
    sheets = list(sheets)
    _validate_sheets(sheets)
    styles, lookup = _styles()
    counts = {}
    with zipfile.ZipFile(target, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=1) as archive:
        types = [XML_HEAD + f'<Types xmlns="{CONTENT_NS}">',
                 '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
                 '<Default Extension="xml" ContentType="application/xml"/>',
                 '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
                 '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>']
        types.extend(f'<Override PartName="/xl/worksheets/sheet{i}.xml" '
                     'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                     for i in range(1, len(sheets) + 1))
        archive.writestr("[Content_Types].xml", "".join(types) + '</Types>')
        archive.writestr("_rels/.rels", XML_HEAD + f'<Relationships xmlns="{PKG_REL_NS}">'
                         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
                         'Target="xl/workbook.xml"/></Relationships>')
        rels = [XML_HEAD + f'<Relationships xmlns="{PKG_REL_NS}">']
        tabs = [XML_HEAD + f'<workbook xmlns="{MAIN_NS}" xmlns:r="{DOC_REL_NS}"><sheets>']
        for index, sheet in enumerate(sheets, 1):
            rels.append(f'<Relationship Id="rId{index}" Type="{DOC_REL_NS}/worksheet" '
                        f'Target="worksheets/sheet{index}.xml"/>')
            tabs.append(f'<sheet name="{_xml(sheet["名称"])}" sheetId="{index}" r:id="rId{index}"/>')
        style_id = len(sheets) + 1
        rels.append(f'<Relationship Id="rId{style_id}" Type="{DOC_REL_NS}/styles" Target="styles.xml"/>')
        archive.writestr("xl/_rels/workbook.xml.rels", "".join(rels) + '</Relationships>')
        archive.writestr("xl/workbook.xml", "".join(tabs) + '</sheets></workbook>')
        archive.writestr("xl/styles.xml", styles)
        for index, sheet in enumerate(sheets, 1):
            counts[sheet["名称"]] = _write_sheet(archive, index, sheet, lookup)
    return counts


def _csv_rows(path, headers):
    with path.open("r", encoding="utf-8-sig", newline="") as source:
        reader = csv.reader(source)
        actual = next(reader, None)
        if actual != headers:
            raise ValueError(f"{path} 表头与清单不一致")
        yield from reader


def _load_manifest(path):
    manifest = json.loads(path.read_text(encoding="utf-8-sig"))
    base = path.parent.resolve()
    sheets = []
    for spec in manifest["工作表"]:
        source = (base / spec["数据文件"]).resolve()
        if source == base or base not in source.parents or source.suffix.lower() != ".csv":
            raise ValueError(f"数据文件须为清单同目录内的 CSV：{spec['数据文件']}")
        if not source.is_file():
            raise FileNotFoundError(source)
        sheets.append({"名称": spec["名称"], "列": spec["列"],
                       "行": _csv_rows(source, [col["名称"] for col in spec["列"]])})
    return sheets


def main():
    parser = argparse.ArgumentParser(description="把已核定的 CSV 明细一次写成深海蓝 Excel")
    parser.add_argument("清单", type=Path, help="UTF-8 JSON 清单路径")
    parser.add_argument("输出", type=Path, help="新建的 XLSX 输出路径")
    args = parser.parse_args()
    target = args.输出.resolve()
    if target.suffix.lower() != ".xlsx" or target.exists():
        parser.error("输出必须是尚不存在的 .xlsx 文件")
    sheets = _load_manifest(args.清单.resolve())
    target.parent.mkdir(parents=True, exist_ok=True)
    partial = target.with_name(target.name + f".partial-{uuid.uuid4().hex}")
    counts = write_workbook(sheets, partial)
    os.rename(partial, target)
    print(json.dumps({"输出": str(target), "数据行数": counts}, ensure_ascii=False))


if __name__ == "__main__":
    main()
