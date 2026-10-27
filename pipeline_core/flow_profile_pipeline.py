"""
flow_profile_pipeline.py

Creates a "Flow Profile" sheet with Time-band rows (06:00-22:00 window)
and per (From -> To) averaged flows based on the processed workbook produced
by `Final Calculation Flow.py`.

Usage:
    python flow_profile_pipeline.py --input Raw Data 1.xlsx

If no input is provided the script defaults to `Raw Data 1.xlsx` in the
current working directory and writes an output workbook named
`<input>_flow_profile.xlsx`.
"""

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import Border, Side, PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter, column_index_from_string
from datetime import time
import math
import argparse
import os

HEADER = [
    "Hour", "Motorcycles", "3-Wheelers/ Tuktuk", "Cars/Taxis/Jeeps/SUVs/Pick-ups",
    "Microbus", "Minibus", "Bus", "Omnibus", "LGV", "MGV", "HGV",
    "Articulated HGV", "Others"
]


def normalize(v):
    return "" if v is None else str(v).strip().lower()


def road_key(v):
    return ''.join(ch for ch in normalize(v) if ch.isalnum())


def format_percent_sigfig(value, sigfigs=4):
    value = to_num(value)
    if value == 0:
        return '0%'
    percent_value = value * 100
    magnitude = int(math.floor(math.log10(abs(percent_value)))) if percent_value != 0 else 0
    decimals = max(sigfigs - magnitude - 1, 0)
    text = f"{percent_value:.{decimals}f}"
    if '.' in text:
        text = text.rstrip('0').rstrip('.')
    return f"{text}%"


def round_sigfigs(value, sigfigs=4):
    value = to_num(value)
    if value == 0:
        return 0.0
    magnitude = math.floor(math.log10(abs(value)))
    decimals = sigfigs - int(magnitude) - 1
    return round(value, decimals)


def to_num(v):
    try:
        if v is None:
            return 0.0
        if isinstance(v, str):
            s = v.strip()
            if s == '':
                return 0.0
            return float(s)
        return float(v)
    except Exception:
        return 0.0


def parse_time(v):
    from datetime import datetime
    if v is None:
        return None
    if isinstance(v, float) and 0.0 <= v < 1.0:
        total_minutes = int(round(v * 24 * 60))
        h, m = divmod(total_minutes, 60)
        return time(h % 24, m)
    if isinstance(v, datetime):
        return v.time()
    if isinstance(v, time):
        return v
    s = str(v).strip()
    for fmt in ("%I:%M %p", "%H:%M", "%I:%M%p", "%I:%M"):
        try:
            return datetime.strptime(s, fmt).time()
        except Exception:
            pass
    return None


def find_headers(ws):
    hits = []
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column - len(HEADER) + 2):
            if all(normalize(ws.cell(r, c + i).value) == normalize(HEADER[i])
                   for i in range(len(HEADER))):
                hits.append((r, c))
    hits.sort(key=lambda x: (x[0], x[1]))
    return hits


def data_rows_of(ws, hdr_row, start_col):
    rows = []
    for r in range(hdr_row + 1, ws.max_row + 1):
        ht = normalize(ws.cell(r, start_col).value)
        if not ht:
            continue
        if ht in ('total', 'max', 'hour', 'direction:', 'date:', 'station:', 'from:', 'to'):
            break
        rows.append(r)
    return rows


def row_totals_of(ws, rows, vcols):
    def to_num(v):
        try:
            return float(v)
        except Exception:
            return 0.0
    return {r: sum(to_num(ws.cell(r, c).value) for c in vcols) for r in rows}


def four_row_blocks(ws, rows, start_col, rtot):
    blocks = []
    for i in range(0, len(rows), 4):
        br = rows[i:i+4]
        if not br:
            continue
        blocks.append({
            'first_row': br[0],
            'last_row':  br[-1],
            'peak':      max(rtot[r] for r in br) * 4,
            'hourly':    sum(rtot[r] for r in br),
            'time':      parse_time(ws.cell(br[0], start_col).value),
        })
    return blocks


def find_direction(ws, header_row, start_col):
    META = {"from:", "from", "to:", "to", "direction:", "date:", "station:", "hour", "total", "max"}
    for r in range(header_row - 1, max(header_row - 10, 0), -1):
        for c in range(start_col, start_col + 20):
            if normalize(ws.cell(r, c).value) in ("from:", "from"):
                for cc in range(c + 1, c + 10):
                    v = ws.cell(r, cc).value
                    nv = normalize(v)
                    if nv and nv not in META:
                        return str(v).strip()
    return 'Unknown'


def find_to(ws, header_row, start_col):
    META = {"from:", "from", "to:", "to:", "direction:", "date:", "station:", "hour", "total", "max"}
    for r in range(header_row - 1, max(header_row - 10, 0), -1):
        for c in range(start_col, start_col + 20):
            if normalize(ws.cell(r, c).value) in ("to:", "to"):
                for cc in range(c + 1, c + 10):
                    v = ws.cell(r, cc).value
                    nv = normalize(v)
                    if nv and nv not in META:
                        return str(v).strip()
    return 'Unknown'


def build_time_slots(start_h=6, start_m=0, end_h=21, end_m=45):
    slots = []
    h, m = start_h, start_m
    while (h < end_h) or (h == end_h and m <= end_m):
        slots.append(time(h, m))
        m += 15
        if m >= 60:
            m -= 60
            h += 1
    return slots


def find_avg_row(ws, header_row, start_col):
    for r in range(header_row + 1, ws.max_row + 1):
        if normalize(ws.cell(r, start_col).value) == '3-day average':
            return r
    return None


thin_side = Side(style='thin', color='7F7F7F')
medium_side = Side(style='medium', color='404040')
thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
medium_border = Border(left=medium_side, right=medium_side, top=medium_side, bottom=medium_side)

title_fill = PatternFill('solid', fgColor='F4F6FA')
header_fill = PatternFill('solid', fgColor='D9EAF7')
subheader_fill = PatternFill('solid', fgColor='EEF4FB')
data_fill = PatternFill(fill_type=None)


def style_range(ws, min_row, max_row, min_col, max_col, *, fill=None, font=None,
                border=None, alignment=None):
    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            if fill is not None:
                cell.fill = fill
            if font is not None:
                cell.font = font
            if border is not None:
                cell.border = border
            if alignment is not None:
                cell.alignment = alignment


def set_formula_text(cell, formula, cached_value=None):
    cell.value = formula


def sheet_ref(sheet_name, coordinate):
    return f"'{sheet_name}'!{coordinate}" if ' ' in sheet_name else f"{sheet_name}!{coordinate}"


def main():
    p = argparse.ArgumentParser()
    p.add_argument('--input', '-i', default='Raw Data 1.xlsx')
    p.add_argument('--processed', '-p', default=None,
                   help='Processed workbook produced by Final Calculation Flow.py')
    p.add_argument('--output', '-o', default=None)
    args = p.parse_args()

    in_path = args.input
    if not os.path.exists(in_path):
        raise FileNotFoundError(f"Input not found: {in_path}")

    def _latest_processed_file():
        candidates = [
            f for f in os.listdir(os.getcwd())
            if f.lower().endswith('.xlsx')
            and '_processed_' in f.lower()
            and not os.path.basename(f).startswith('~$')
        ]
        if not candidates:
            return None
        candidates.sort(key=lambda name: os.path.getmtime(os.path.join(os.getcwd(), name)), reverse=True)
        return os.path.join(os.getcwd(), candidates[0])

    source_path = args.processed or _latest_processed_file() or in_path
    if not os.path.exists(source_path):
        raise FileNotFoundError(f"Source workbook not found: {source_path}")

    out_path = args.output or os.path.splitext(source_path)[0] + '_flow_profile.xlsx'

    wb = load_workbook(source_path, data_only=True)

    header_positions = []
    ws = None
    for candidate in wb.worksheets:
        hits = find_headers(candidate)
        if hits:
            ws = candidate
            header_positions = hits
            break
    if ws is None or not header_positions:
        raise ValueError('No HEADER blocks found in workbook')

    # Build block registry similar to Final Calculation Flow
    block_registry = []
    time_slots = build_time_slots(6,0,21,45)

    for hdr_row, sc in header_positions:
        vcols = [sc + i for i in range(1, len(HEADER))]
        drows = data_rows_of(ws, hdr_row, sc)
        if not drows:
            continue
        rtot = row_totals_of(ws, drows, vcols)
        # Build cell_time_map
        time_map = {}
        cur_t = None
        for r in range(hdr_row + 1, ws.max_row + 1):
            cell_val = ws.cell(r, sc).value
            norm_val = normalize(cell_val)
            if norm_val in ('total', 'max', 'hour', 'direction:', 'date:', 'station:', 'from:', 'to'):
                break
            pt = parse_time(cell_val)
            if pt is not None:
                cur_t = pt
            elif cur_t is not None:
                mins = cur_t.hour * 60 + cur_t.minute + 15
                cur_t = time(mins // 60 % 24, mins % 60)
            if cur_t is None:
                continue
            row_val = sum(to_num(ws.cell(r, c).value) for c in vcols)
            time_map.setdefault(cur_t, []).append((r, sc + len(HEADER), row_val))

        block_registry.append({
            'start_col': sc,
            'header_row': hdr_row,
            'from_road': find_direction(ws, hdr_row, sc),
            'to_road': find_to(ws, hdr_row, sc),
            'cell_time_map': time_map,
            'avg_row': find_avg_row(ws, hdr_row, sc),
        })

    # Build ordered unique (from,to) pairs
    pair_keys = []
    seen = set()
    for b in block_registry:
        key = (b['from_road'], b['to_road'])
        if key in seen:
            continue
        seen.add(key)
        pair_keys.append(key)

    # Create new workbook or sheet for Flow Profile
    # We'll add sheet to original workbook and save as a new file so original is untouched
    if 'Flow Profile' in wb.sheetnames:
        del wb['Flow Profile']
        wsf = wb.create_sheet('Flow Profile')
    else:
        wsf = wb.create_sheet('Flow Profile')

    # Original 15-min layout: title row, then one header row and time rows below.
    total_cols = 1 + len(pair_keys)
    wsf.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_cols)
    title_cell = wsf.cell(1, 1, value='15-min Diurnal Variation in Traffic Flow')
    try:
        from openpyxl.styles import Font, Alignment
        title_cell.font = Font(bold=True, color='FF0000')
        title_cell.alignment = Alignment(horizontal='left', vertical='center')
    except Exception:
        pass

    wsf.cell(2, 1, value='Time band')
    try:
        from openpyxl.styles import Font, Alignment
        wsf.cell(2, 1).font = Font(bold=True)
        wsf.cell(2, 1).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    except Exception:
        pass
    for idx, (fr, to) in enumerate(pair_keys, start=2):
        c = wsf.cell(2, idx, value=f'{fr} {to}')
        try:
            c.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
            c.font = Font(bold=True)
        except Exception:
            pass

    from datetime import datetime, timedelta
    for r_off, t in enumerate(time_slots, start=3):
        start_dt = datetime.combine(datetime.today(), t)
        end_dt = start_dt + timedelta(minutes=15)
        cell = wsf.cell(r_off, 1, value=f"{start_dt.strftime('%H:%M')}-{end_dt.strftime('%H:%M')}")
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = thin_border
        for c_off, (fr, to) in enumerate(pair_keys, start=2):
            vals = []
            for b in block_registry:
                if b['from_road'] == fr and b['to_road'] == to:
                    entries = b.get('cell_time_map', {}).get(t, [])
                    for _row, _tc, val in entries:
                        vals.append(val)
            data_cell = wsf.cell(r_off, c_off, value=round(sum(vals) / len(vals), 0) if vals else 0)
            data_cell.alignment = Alignment(horizontal='center', vertical='center')
            data_cell.border = thin_border

    # Apply table styling to the 15-minute table
    style_range(
        wsf,
        1,
        2 + len(time_slots),
        1,
        total_cols,
        fill=None,
        border=thin_border,
        alignment=Alignment(horizontal='center', vertical='center', wrap_text=True),
    )
    wsf.cell(1, 1).fill = title_fill
    wsf.cell(1, 1).font = Font(bold=True, color='C00000', size=12)
    wsf.cell(2, 1).fill = header_fill
    wsf.cell(2, 1).font = Font(bold=True, size=9)
    for idx in range(2, total_cols + 1):
        wsf.cell(2, idx).fill = header_fill
        wsf.cell(2, idx).font = Font(bold=True, size=9)
        wsf.cell(2, idx).border = medium_border
    for r in range(3, 3 + len(time_slots)):
        wsf.cell(r, 1).fill = subheader_fill
        for c in range(2, total_cols + 1):
            wsf.cell(r, c).border = thin_border

    # Second table: 16-hr Traffic Flows, placed 6 blank rows below the first table
    first_table_last_row = 2 + len(time_slots)
    second_title_row = first_table_last_row + 7
    second_header_row = second_title_row + 1
    second_data_start = second_title_row + 2

    second_width = 1 + len(HEADER)  # pair label + 12 vehicle classes + total
    wsf.merge_cells(start_row=second_title_row, start_column=1, end_row=second_title_row, end_column=second_width)
    second_title = wsf.cell(second_title_row, 1, value='16-hr Traffic Flows')
    try:
        from openpyxl.styles import Font, Alignment
        second_title.font = Font(bold=True, color='FF0000')
        second_title.alignment = Alignment(horizontal='left', vertical='center')
    except Exception:
        pass

    try:
        from openpyxl.styles import Font, Alignment
        label_cell = wsf.cell(second_header_row, 1, value='From / To')
        label_cell.font = Font(bold=True)
        label_cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        for idx, col_name in enumerate(HEADER[1:], start=2):
            hdr_cell = wsf.cell(second_header_row, idx, value=col_name)
            hdr_cell.font = Font(bold=True)
            hdr_cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    except Exception:
        wsf.cell(second_header_row, 1, value='From / To')
        for idx, col_name in enumerate(HEADER[1:], start=2):
            wsf.cell(second_header_row, idx, value=col_name)

    # Build one row per intelligent road pair from the 3-day average row.
    # The source row already contains the average, so do not divide again.
    row_data_by_pair = {}
    for b in block_registry:
        avg_row = b.get('avg_row')
        if avg_row is None:
            continue
        key = (b['from_road'], b['to_road'])
        values = []
        for offset in range(1, len(HEADER)):
            values.append(to_num(ws.cell(avg_row, b['start_col'] + offset).value))
        total_val = to_num(ws.cell(avg_row, b['start_col'] + len(HEADER)).value)
        if key not in row_data_by_pair:
            row_data_by_pair[key] = {
                'values': values,
                'total': total_val,
            }
        else:
            # Keep the first observed 3-day average row for the pair.
            # This avoids averaging an already-averaged row down to zero.
            continue

    for row_idx, (fr, to) in enumerate(pair_keys, start=second_data_start):
        data = row_data_by_pair.get((fr, to))
        wsf.cell(row_idx, 1, value=f'{fr}-{to}')
        try:
            from openpyxl.styles import Alignment
            wsf.cell(row_idx, 1).alignment = Alignment(horizontal='left', vertical='center')
        except Exception:
            pass
        if not data:
            continue
        for col_idx, value in enumerate(data['values'], start=2):
            data_cell = wsf.cell(row_idx, col_idx, value=int(round(value)) if abs(value - round(value)) < 1e-9 else round(value, 0))
            data_cell.alignment = Alignment(horizontal='center', vertical='center')
            data_cell.border = thin_border
        total_cell = wsf.cell(row_idx, second_width)
        total_cell.value = int(round(data['total'])) if abs(data['total'] - round(data['total'])) < 1e-9 else round(data['total'], 0)
        total_cell.alignment = Alignment(horizontal='center', vertical='center')
        total_cell.border = medium_border

    second_row_map = {pair_keys[i]: second_data_start + i for i in range(len(pair_keys))}

    # Apply table styling to the 16-hour table
    style_range(
        wsf,
        second_title_row,
        second_data_start + len(pair_keys) - 1,
        1,
        second_width,
        fill=None,
        border=thin_border,
        alignment=Alignment(horizontal='center', vertical='center', wrap_text=True),
    )
    wsf.cell(second_title_row, 1).fill = title_fill
    wsf.cell(second_title_row, 1).font = Font(bold=True, color='C00000', size=12)
    for c in range(1, second_width + 1):
        wsf.cell(second_header_row, c).fill = header_fill
        wsf.cell(second_header_row, c).font = Font(bold=True, size=9)
        wsf.cell(second_header_row, c).border = medium_border
    for r in range(second_data_start, second_data_start + len(pair_keys)):
        wsf.cell(r, 1).fill = subheader_fill
        wsf.cell(r, 1).font = Font(bold=True, size=9)
        for c in range(2, second_width + 1):
            wsf.cell(r, c).border = thin_border

    # Percentage table derived from the 16-hr table.
    percent_title_row = second_data_start + len(pair_keys) + 3
    percent_header_row = percent_title_row + 1
    percent_data_start = percent_title_row + 2
    percent_width = 10

    wsf.merge_cells(start_row=percent_title_row, start_column=1, end_row=percent_title_row, end_column=percent_width)
    percent_title = wsf.cell(percent_title_row, 1, value='Percentage Traffic Flows')
    percent_title.fill = title_fill
    percent_title.font = Font(bold=True, color='C00000', size=12)
    percent_title.alignment = Alignment(horizontal='left', vertical='center')
    percent_title.border = thin_border

    percent_headers = [
        'From / To',
        '% of Micromobility',
        '',
        'Cars/Taxis/Jeeps/SUVs/Pick-ups / Total',
        'Microbus+Minibus+Bus+Omnibus / Total',
        '',
        '',
        '',
        'LGV / Total',
        'MGV+HGV+Articulated HGV+Others / Total',
    ]

    for idx, header_text in enumerate(percent_headers, start=1):
        hdr = wsf.cell(percent_header_row, idx, value=header_text)
        hdr.fill = header_fill
        hdr.font = Font(bold=True, size=9)
        hdr.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        hdr.border = medium_border

    percent_row_refs = {}
    for row_idx, pair in enumerate(pair_keys, start=percent_data_start):
        src_row = second_row_map.get(pair)
        if src_row is None:
            continue
        fr, to = pair
        percent_row_refs[pair] = row_idx
        wsf.cell(row_idx, 1, value=f'{pair[0]}-{pair[1]}')
        wsf.cell(row_idx, 1).alignment = Alignment(horizontal='left', vertical='center')
        wsf.cell(row_idx, 1).border = thin_border

        is_same_road = road_key(fr) == road_key(to)

        # 1st cell on the right: Micromobility = (Motorcycles + 3-Wheelers/Tuktuk) / Total
        micromobility = 0
        total_val = to_num(wsf.cell(src_row, second_width).value)
        if total_val and not is_same_road:
            micromobility = (to_num(wsf.cell(src_row, 2).value) + to_num(wsf.cell(src_row, 3).value)) / total_val

        # 3rd and 4th cells on the right
        cars_share = 0
        mixed_share = 0
        if total_val and not is_same_road:
            cars_share = to_num(wsf.cell(src_row, 4).value) / total_val
            mixed_share = (
                to_num(wsf.cell(src_row, 5).value) +
                to_num(wsf.cell(src_row, 6).value) +
                to_num(wsf.cell(src_row, 7).value) +
                to_num(wsf.cell(src_row, 8).value)
            ) / total_val

        # 8th and 9th cells on the right
        lgv_share = 0
        heavy_share = 0
        if total_val and not is_same_road:
            lgv_share = to_num(wsf.cell(src_row, 9).value) / total_val
            heavy_share = (
                to_num(wsf.cell(src_row, 10).value) +
                to_num(wsf.cell(src_row, 11).value) +
                to_num(wsf.cell(src_row, 12).value) +
                to_num(wsf.cell(src_row, 13).value)
            ) / total_val

        percent_values = {
            2: micromobility,
            4: cars_share,
            5: mixed_share,
            9: lgv_share,
            10: heavy_share,
        }

        for col_idx in range(2, percent_width + 1):
            cell = wsf.cell(row_idx, col_idx)
            if col_idx in percent_values:
                cell.value = percent_values[col_idx]
                cell.number_format = '0.0000%'
            else:
                cell.value = ''
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = thin_border

    style_range(
        wsf,
        percent_title_row,
        percent_data_start + len(pair_keys) - 1,
        1,
        percent_width,
        fill=None,
        border=thin_border,
        alignment=Alignment(horizontal='center', vertical='center', wrap_text=True),
    )
    wsf.cell(percent_title_row, 1).fill = title_fill
    wsf.cell(percent_title_row, 1).font = Font(bold=True, color='C00000', size=12)

    for c in range(1, percent_width + 1):
        wsf.cell(percent_header_row, c).fill = header_fill
        wsf.cell(percent_header_row, c).font = Font(bold=True, size=9)
        wsf.cell(percent_header_row, c).border = medium_border

    for r in range(percent_data_start, percent_data_start + len(pair_keys)):
        wsf.cell(r, 1).fill = subheader_fill
        wsf.cell(r, 1).font = Font(bold=True, size=9)
        for c in range(2, percent_width + 1):
            wsf.cell(r, c).border = thin_border

    average_percent_row = percent_data_start + len(pair_keys)
    avg_label = wsf.cell(average_percent_row, 1, value='Average Percentage')
    avg_label.fill = PatternFill('solid', fgColor='C00000')
    avg_label.font = Font(bold=True, color='FFFFFF', size=9)
    avg_label.alignment = Alignment(horizontal='center', vertical='center')
    avg_label.border = medium_border

    for col_idx in range(2, percent_width + 1):
        values = [to_num(wsf.cell(r, col_idx).value) for r in range(percent_data_start, percent_data_start + len(pair_keys))]
        values = [v for v in values if v != 0]
        avg_cell = wsf.cell(average_percent_row, col_idx)
        avg_cell.fill = PatternFill('solid', fgColor='C00000')
        avg_cell.font = Font(bold=True, color='FFFFFF', size=9)
        avg_cell.alignment = Alignment(horizontal='center', vertical='center')
        avg_cell.border = medium_border
        if values:
            avg_cell.value = sum(values) / len(values)
            avg_cell.number_format = '0.0000%'
        else:
            avg_cell.value = ''

    wsf.row_dimensions[average_percent_row].height = 22

    # Hourly Diurnal Variation table based on sums of the 15-minute table.
    # Place it at the top of the sheet, to the right of the main 15-minute table.
    hourly_title_row = 1
    hourly_header_row = hourly_title_row + 1
    hourly_data_start = hourly_title_row + 2
    hourly_width = total_cols
    chart_anchor_col = total_cols + 3
    hourly_start_col = chart_anchor_col + 10

    wsf.merge_cells(
        start_row=hourly_title_row,
        start_column=hourly_start_col,
        end_row=hourly_title_row,
        end_column=hourly_start_col + hourly_width - 1,
    )
    hourly_title = wsf.cell(hourly_title_row, hourly_start_col, value='Hourly Diurnal Variation in Traffic Flow')
    hourly_title.fill = title_fill
    hourly_title.font = Font(bold=True, color='C00000', size=12)
    hourly_title.alignment = Alignment(horizontal='left', vertical='center')
    hourly_title.border = thin_border

    hourly_head = wsf.cell(hourly_header_row, hourly_start_col, value='Time band')
    hourly_head.fill = header_fill
    hourly_head.font = Font(bold=True, size=9)
    hourly_head.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    hourly_head.border = medium_border

    for idx, (fr, to) in enumerate(pair_keys, start=2):
        hdr = wsf.cell(hourly_header_row, hourly_start_col + idx - 1, value=f'{fr} {to}')
        hdr.fill = header_fill
        hdr.font = Font(bold=True, size=9)
        hdr.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        hdr.border = medium_border

    for hour_idx in range(16):
        start_hour = 6 + hour_idx
        end_hour = start_hour + 1
        row_idx = hourly_data_start + hour_idx
        label_cell = wsf.cell(row_idx, hourly_start_col, value=f'{start_hour:02d}:00-{end_hour:02d}:00')
        label_cell.fill = subheader_fill
        label_cell.font = Font(bold=True, size=9)
        label_cell.alignment = Alignment(horizontal='center', vertical='center')
        label_cell.border = thin_border

        start_15min_row = 3 + (hour_idx * 4)
        end_15min_row = start_15min_row + 3
        for col_idx in range(2, total_cols + 1):
            hour_sum = 0
            for r in range(start_15min_row, end_15min_row + 1):
                hour_sum += to_num(wsf.cell(r, col_idx).value)
            out_val = int(round(hour_sum)) if abs(hour_sum - round(hour_sum)) < 1e-9 else round(hour_sum, 0)
            data_cell = wsf.cell(row_idx, hourly_start_col + col_idx - 1, value=out_val)
            data_cell.alignment = Alignment(horizontal='center', vertical='center')
            data_cell.border = thin_border

    style_range(
        wsf,
        hourly_title_row,
        hourly_data_start + 15,
        hourly_start_col,
        hourly_start_col + hourly_width - 1,
        fill=None,
        border=thin_border,
        alignment=Alignment(horizontal='center', vertical='center', wrap_text=True),
    )

    # Narrow column widths to roughly 40 pixels and keep headers wrapped.
    # Reduce widths by ~30% from the previous compact layout.
    # Excel column width is not pixels, so this is an approximate visual match.
    compact_width = 4.8
    first_col_width = 5.6

    for col_idx in range(1, total_cols + 1):
        wsf.column_dimensions[get_column_letter(col_idx)].width = compact_width
    for col_idx in range(1, second_width + 1):
        wsf.column_dimensions[get_column_letter(col_idx)].width = compact_width

    wsf.column_dimensions[get_column_letter(1)].width = first_col_width
    wsf.row_dimensions[1].height = 30
    wsf.row_dimensions[2].height = 50
    wsf.row_dimensions[second_title_row].height = 30
    wsf.row_dimensions[second_header_row].height = 50
    wsf.row_dimensions[hourly_title_row].height = 30
    wsf.row_dimensions[hourly_header_row].height = 50
    for r in range(3, 3 + len(time_slots)):
        wsf.row_dimensions[r].height = 20
    for r in range(second_data_start, second_data_start + len(pair_keys)):
        wsf.row_dimensions[r].height = 25
    for r in range(hourly_data_start, hourly_data_start + 16):
        wsf.row_dimensions[r].height = 20

    # Line chart for the 15-minute traffic flow table
    if pair_keys:
        chart = LineChart()
        chart.title = '15-min Traffic Flow by Road Pair'
        chart.style = 13
        chart.y_axis.title = 'Traffic Flow'
        chart.x_axis.title = 'Time'
        chart.height = 13
        chart.width = 30
        chart.y_axis.majorGridlines = None
        chart.y_axis.delete = False
        chart.x_axis.delete = False
        chart.y_axis.spPr = None
        chart.x_axis.spPr = None
        try:
            chart.legend = None
        except Exception:
            pass

        data = Reference(wsf, min_col=2, max_col=1 + len(pair_keys), min_row=2, max_row=2 + len(time_slots))
        cats = Reference(wsf, min_col=1, min_row=3, max_row=2 + len(time_slots))
        chart.add_data(data, titles_from_data=True, from_rows=False)
        chart.set_categories(cats)
        chart.scatterStyle = None

        palette = [
            '1f77b4', 'ff7f0e', '2ca02c', 'd62728', '9467bd', '8c564b',
            'e377c2', '7f7f7f', 'bcbd22', '17becf', '003f5c', 'ffa600',
            '665191', '2f4b7c', 'f95d6a', 'ff6f61'
        ]
        for idx, series in enumerate(chart.series):
            color = palette[idx % len(palette)]
            try:
                series.graphicalProperties.line.solidFill = color
                series.graphicalProperties.line.width = 6350
                series.graphicalProperties.solidFill = color
            except Exception:
                pass

        chart_anchor_row = 3
        wsf.add_chart(chart, f'{get_column_letter(chart_anchor_col)}{chart_anchor_row}')

    try:
        wsf.sheet_view.showFormulas = True
    except Exception:
        pass

    wb.save(out_path)
    print(f'Flow Profile written to: {out_path}')


if __name__ == '__main__':
    main()
