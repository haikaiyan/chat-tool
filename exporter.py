"""导出：在原表基础上新增「是否命中」「命中的关键词」两列，原文件不修改。"""
from typing import List

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill

from scanner import CONTENT_HEADERS, HitRecord, _norm, _find_col

HIT_COL = "是否命中"
KW_COL = "命中的关键词"

HEADER_FILL = PatternFill("solid", fgColor="DDEBF7")
HIT_FILL = PatternFill("solid", fgColor="FFF2CC")


def _content_col(headers: List[str], fallback_index: int = None, ws=None):
    idx = _find_col(headers, CONTENT_HEADERS)
    if idx is not None:
        return idx
    if fallback_index is not None and ws is not None:
        return fallback_index
    return None


def _last_used_col(headers: List[str], ws) -> int:
    """返回最后一列的 1-based 列号。"""
    last = 0
    for i, h in enumerate(headers, start=1):
        if h:
            last = i
    if last == 0:
        last = ws.max_column or 1
    return last


def export_marked(src_path: str, dest_path: str, hits: List[HitRecord]) -> int:
    """生成标记版 Excel：每个被扫描的工作表末尾追加 是否命中 / 命中的关键词 两列。"""
    hit_map = {}
    for h in hits:
        hit_map.setdefault(h.sheet, {})[h.row] = h

    wb = load_workbook(src_path)
    marked_rows = 0
    for ws in wb.worksheets:
        if ws.title not in hit_map:
            continue
        headers = [_norm(c.value) for c in ws[1]]
        start = _last_used_col(headers, ws) + 1

        c1 = ws.cell(row=1, column=start, value=HIT_COL)
        c2 = ws.cell(row=1, column=start + 1, value=KW_COL)
        for c in (c1, c2):
            c.font = Font(bold=True)
            c.fill = HEADER_FILL

        content_idx = _find_col(headers, CONTENT_HEADERS)
        sheet_hits = hit_map[ws.title]
        for row in range(2, ws.max_row + 1):
            # 空行不标记，避免给 2 万多行空白行写入“否”
            if content_idx is not None:
                v = ws.cell(row=row, column=content_idx + 1).value
                if v is None or str(v).strip() == "":
                    continue
            h = sheet_hits.get(row)
            if h:
                ws.cell(row=row, column=start, value="是").fill = HIT_FILL
                ws.cell(row=row, column=start + 1, value="、".join(h.keywords))
                marked_rows += 1
            else:
                ws.cell(row=row, column=start, value="否")
                ws.cell(row=row, column=start + 1, value="")

    wb.save(dest_path)
    return marked_rows


def export_hits_only(src_path: str, dest_path: str, hits: List[HitRecord]) -> int:
    """仅导出命中行（保留原表列 + 新增两列），便于分发复核。"""
    by_sheet = {}
    for h in hits:
        by_sheet.setdefault(h.sheet, []).append(h)

    src = load_workbook(src_path, read_only=True, data_only=True)
    out = Workbook()
    out.remove(out.active)
    total = 0

    for ws in src.worksheets:
        if ws.title not in by_sheet:
            continue
        rows_iter = ws.iter_rows(values_only=True)
        header = next(rows_iter, None)
        if header is None:
            continue
        header = list(header)
        row_map = {}
        wanted = {h.row for h in by_sheet[ws.title]}
        for r_idx, row in enumerate(rows_iter, start=2):
            if r_idx in wanted:
                row_map[r_idx] = list(row)
            elif wanted and r_idx > max(wanted):
                break

        sheet = out.create_sheet(title=ws.title[:31])
        for i, v in enumerate(header, start=1):
            c = sheet.cell(row=1, column=i, value=v)
            c.font = Font(bold=True)
            c.fill = HEADER_FILL
        start = len(header) + 1
        c1 = sheet.cell(row=1, column=start, value=HIT_COL)
        c2 = sheet.cell(row=1, column=start + 1, value=KW_COL)
        for c in (c1, c2):
            c.font = Font(bold=True)
            c.fill = HEADER_FILL

        for r, h in enumerate(sorted(by_sheet[ws.title], key=lambda x: x.row), start=2):
            values = row_map.get(h.row, [])
            for i, v in enumerate(values, start=1):
                sheet.cell(row=r, column=i, value=v)
            sheet.cell(row=r, column=start, value="是").fill = HIT_FILL
            sheet.cell(row=r, column=start + 1, value="、".join(h.keywords))
            total += 1

    src.close()
    out.save(dest_path)
    return total
