"""
Traffic count processor
=======================
PIPELINE
  Stage 1 – Parse all source files from raw vehicle counts.
             Nothing is written yet; all results are held in memory.
  Stage 2 – Open the target workbook and write everything at once.
             Blocks are grouped by from-road; each (from-road, day) group
             gets its own summary table placed immediately after its last
             block. Blocks from subsequent from-road groups are relocated
             if a summary table would otherwise land on top of them.
"""

from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Alignment, Font, Border, Side
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.cell.cell import MergedCell
import os, re
from collections import defaultdict
from datetime import datetime, time

# Patch openpyxl so formula cells also write a <v> cached value, which lets
# Excel display the number immediately without needing to recalculate.
# Usage: cell._cached_value = <number>  after  cell.value = "=formula"
import openpyxl.cell._writer as _opxl_cw
from openpyxl.xml.functions import Element, SubElement
from openpyxl.compat import safe_string as _safe_string

# Dict mapping cell coordinate string (e.g. "N7") to pre-computed numeric value.
# The patched writer uses this to emit <v>value</v> alongside <f>formula</f>
# so Excel displays results immediately without needing to recalculate.
_formula_cache: dict = {}

_orig_write_cell = _opxl_cw.etree_write_cell

def _write_cell_with_cache(xf, worksheet, cell, styled=None):
    cached = _formula_cache.get(cell.coordinate)
    if cell.data_type == 'f' and cached is not None:
        value, attributes = _opxl_cw._set_attributes(cell, styled)
        el = Element("c", attributes)
        if value is None or value == "":
            xf.write(el)
            return
        formula = SubElement(el, 'f')
        formula.text = value[1:]
        v_el = SubElement(el, 'v')
        v_el.text = _safe_string(cached)
        xf.write(el)
    else:
        _orig_write_cell(xf, worksheet, cell, styled)

_opxl_cw.etree_write_cell = _write_cell_with_cache
_opxl_cw.write_cell       = _write_cell_with_cache

import openpyxl.worksheet._writer as _opxl_ws_writer
_opxl_ws_writer.write_cell = _write_cell_with_cache


def _set_cached(ws, row, col, value):
    """Record a pre-computed value for a formula cell (keyed by coordinate)."""
    cell = ws.cell(row, col)
    if isinstance(cell, MergedCell):
        for rng in ws.merged_cells.ranges:
            if cell.coordinate in rng:
                cell = ws.cell(rng.min_row, rng.min_col)
                break
    _formula_cache[cell.coordinate] = value

# ── configuration ──────────────────────────────────────────────────────────

file_path   = r"Raw Data 1.xlsx"  # comma-separated list of input files
output_path = file_path[:-5] + f"_processed_{datetime.now():%Y%m%d_%H%M%S}.xlsx"

FILES = [
    file_path
]

HEADER = [
    "Hour", "Motorcycles", "3-Wheelers/ Tuktuk", "Cars/Taxis/Jeeps/SUVs/Pick-ups",
    "Microbus", "Minibus", "Bus", "Omnibus", "LGV", "MGV", "HGV",
    "Articulated HGV", "Others"
]

PERIODS = {
    'morning': (6,  0,  9, 45),
    'lunch':   (11, 0, 14, 45),
    'evening': (16, 0, 19, 45),
}

COMPUTED_COLS = 3                           # TOTAL, Peak Hourly Flow, Hourly Flow
BLOCK_WIDTH   = len(HEADER) + COMPUTED_COLS # actual output columns per block (16)
BLOCK_STRIDE  = BLOCK_WIDTH + 1            # column advance between block starts (17)
SUMMARY_GAP   = 1   # blank cols between last block of group and its summary table
SUMMARY_TRAIL = 2   # blank cols between summary table end and next group start

_no_fill   = PatternFill(fill_type=None)
_no_border = Border()
_no_font   = Font()

# ── helpers ────────────────────────────────────────────────────────────────

def normalize(v):
    return "" if v is None else str(v).strip().lower()

def to_num(v):
    try:    return float(v)
    except: return 0.0

def parse_time(v):
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
        try:    return datetime.strptime(s, fmt).time()
        except: pass
    return None

def in_period(t, sh, sm, eh, em):
    return t is not None and time(sh, sm) <= t <= time(eh, em)

def find_headers(ws):
    hits = []
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column - len(HEADER) + 2):
            if all(normalize(ws.cell(r, c+i).value) == normalize(HEADER[i])
                   for i in range(len(HEADER))):
                hits.append((r, c))
    hits.sort(key=lambda x: (x[0], x[1]))
    return hits

def data_rows_of(ws, hdr_row, start_col):
    rows = []
    for r in range(hdr_row + 1, ws.max_row + 1):
        ht = normalize(ws.cell(r, start_col).value)
        if not ht:      continue
        if ht == 'total': break
        if ht == 'hour':  break
        if ht in ('direction:', 'date:', 'station:', 'from:', 'to') or ht.startswith('direc'):
            break
        rows.append(r)
    return rows

def row_totals_of(ws, rows, vcols):
    return {r: sum(to_num(ws.cell(r, c).value) for c in vcols) for r in rows}

def four_row_blocks(ws, rows, start_col, rtot):
    blocks = []
    for i in range(0, len(rows), 4):
        br = rows[i:i+4]
        if not br: continue
        blocks.append({
            'first_row': br[0],
            'last_row':  br[-1],
            'peak':      max(rtot[r] for r in br) * 4,
            'hourly':    sum(rtot[r] for r in br),
            'time':      parse_time(ws.cell(br[0], start_col).value),
        })
    return blocks

def safe_set(ws, row, col, value):
    cell = ws.cell(row, col)
    if isinstance(cell, MergedCell):
        for rng in ws.merged_cells.ranges:
            if cell.coordinate in rng:
                ws.cell(rng.min_row, rng.min_col, value=value)
                return
    else:
        cell.value = value

def safe_style(ws, row, col, *, fill=None, alignment=None, font=None,
               border=None, number_format=None):
    cell = ws.cell(row, col)
    if isinstance(cell, MergedCell):
        for rng in ws.merged_cells.ranges:
            if cell.coordinate in rng:
                cell = ws.cell(rng.min_row, rng.min_col)
                break
    if fill          is not None: cell.fill          = fill
    if alignment     is not None: cell.alignment     = alignment
    if font          is not None: cell.font          = font
    if border        is not None: cell.border        = border
    if number_format is not None: cell.number_format = number_format

def _resolve_cell(ws, row, col, _depth=0):
    """Return cell value, following =REF chains up to 5 levels deep."""
    v = ws.cell(row, col).value
    if _depth < 5 and isinstance(v, str) and v.startswith('='):
        m = re.match(r'^=([A-Z]+)(\d+)$', v.strip())
        if m:
            v = _resolve_cell(ws,
                              int(m.group(2)),
                              column_index_from_string(m.group(1)),
                              _depth + 1)
    return v

_META_LABELS = {"from:", "from", "to:", "to", "direction:", "date:", "station:", "hour", "total", "max"}

def find_direction(ws, header_row, start_col):
    for r in range(header_row - 1, max(header_row - 10, 0), -1):
        for c in range(start_col, start_col + BLOCK_STRIDE):
            if normalize(ws.cell(r, c).value) in ("from:", "from"):
                for cc in range(c + 1, c + 10):
                    v = _resolve_cell(ws, r, cc)
                    nv = normalize(v)
                    if nv and nv not in _META_LABELS:
                        return str(v).strip()
    return "Unknown"

def find_to(ws, header_row, start_col):
    for r in range(header_row - 1, max(header_row - 10, 0), -1):
        for c in range(start_col, start_col + BLOCK_STRIDE):
            if normalize(ws.cell(r, c).value) in ("to:", "to"):
                for cc in range(c + 1, c + 10):
                    v = _resolve_cell(ws, r, cc)
                    nv = normalize(v)
                    if nv and nv not in _META_LABELS:
                        return str(v).strip()
    return "Unknown"

def shift_formula_cols(formula, offset):
    """Shift all column letter references in a formula by offset columns."""
    def replace_col(m):
        new_col = column_index_from_string(m.group(1)) + offset
        return get_column_letter(new_col) + m.group(2)
    return re.sub(r'\b([A-Z]{1,3})(\d+)', replace_col, formula)


# ══════════════════════════════════════════════════════════════════════════
# STAGE 1 – Parse all source files; hold results in memory
# ══════════════════════════════════════════════════════════════════════════

col_period_peaks    = {}
col_period_hourlies = {}
col_file_totals     = {}

print("Stage 1 – parsing source files …")

for fpath in FILES:
    if not os.path.exists(fpath):
        print(f"  SKIP (not found): {fpath}")
        continue

    print(f"  reading: {fpath}")
    wb_f = load_workbook(fpath, data_only=True)
    ws_f = wb_f[wb_f.sheetnames[0]]
    positions = find_headers(ws_f)
    if not positions:
        print(f"  SKIP (no header found): {fpath}")
        continue

    col_this_file = {}

    for idx, (hdr, sc) in enumerate(positions):
        vcols  = [sc + i for i in range(1, len(HEADER))]
        drows  = data_rows_of(ws_f, hdr, sc)
        if not drows:
            continue
        rtot   = row_totals_of(ws_f, drows, vcols)
        blocks = four_row_blocks(ws_f, drows, sc, rtot)

        if sc not in col_period_peaks:
            col_period_peaks[sc]    = {p: [] for p in PERIODS}
            col_period_hourlies[sc] = {p: [] for p in PERIODS}

        for pname, (sh, sm, eh, em) in PERIODS.items():
            sel = [b for b in blocks if in_period(b['time'], sh, sm, eh, em)]
            for b in sel:
                col_period_peaks[sc][pname].append(b['peak'])
                col_period_hourlies[sc][pname].append(b['hourly'])

        total_row_r = None
        for r in range(hdr + 1, ws_f.max_row + 1):
            if normalize(ws_f.cell(r, sc).value) == 'max':
                total_row_r = r - 1
                break

        if total_row_r is not None:
            if sc not in col_this_file:
                col_this_file[sc] = {}
            for offset, col in enumerate(vcols, start=1):
                col_this_file[sc][offset] = (
                    col_this_file[sc].get(offset, 0)
                    + to_num(ws_f.cell(total_row_r, col).value)
                )
            col_this_file[sc][13] = (
                col_this_file[sc].get(13, 0)
                + sum(to_num(ws_f.cell(total_row_r, col).value) for col in vcols)
            )

    for sc, totals in col_this_file.items():
        col_file_totals.setdefault(sc, []).append(totals)

print("Stage 1 complete.\n")

# STAGE 2 – Open target workbook and write all results


print("Stage 2 writing output …")

wb = load_workbook(file_path)
ws = wb["Sheet1"]

for tbl_name in list(ws.tables.keys()):
    del ws.tables[tbl_name]

for cd in ws.column_dimensions.values():
    cd.font = _no_font

block_positions = find_headers(ws)
if not block_positions:
    raise ValueError("No header blocks found in Sheet1.")

# ── Layout computation ─────────────────────────────────────────────────────
# Group blocks by (header_row, from_road).  For each row, interleave:
#   [group-A blocks] [group-A summary] [group-B blocks] [group-B summary] …
# Returns:
#   new_start_col  : (hr, orig_sc) -> new_sc
#   summary_start  : (hr, from_road) -> summary start col
#   row_groups     : hr -> [(from_road, [orig_scs], [to_roads]), ...]

def _summary_width(to_roads):
    return 1 + len(to_roads) + 3   # Time + N destinations + TOTAL + MaxHourly + Sum

def compute_layout(block_positions, ws):
    row_blocks = defaultdict(list)
    for hr, sc in block_positions:
        row_blocks[hr].append(sc)

    # Build a fallback from_road per start_col using the first non-Unknown day found
    sc_from_road = {}
    for hr, sc in block_positions:
        if sc not in sc_from_road:
            fr = find_direction(ws, hr, sc)
            if fr != 'Unknown':
                sc_from_road[sc] = fr
    # Second pass: fill remaining 'Unknown' from the fallback map
    for hr, sc in block_positions:
        if sc not in sc_from_road:
            sc_from_road[sc] = 'Unknown'

    # Same fallback for to_road
    sc_to_road = {}
    for hr, sc in block_positions:
        if sc not in sc_to_road:
            to = find_to(ws, hr, sc)
            if to != 'Unknown':
                sc_to_road[sc] = to
    for hr, sc in block_positions:
        if sc not in sc_to_road:
            sc_to_road[sc] = 'Unknown'

    new_start_col = {}
    summary_start = {}
    row_groups    = {}

    for hr in sorted(row_blocks):
        scs = sorted(row_blocks[hr])
        # Use per-day detection where available, fall back to cross-day map
        fr_map = {}
        for sc in scs:
            fr = find_direction(ws, hr, sc)
            fr_map[sc] = fr if fr != 'Unknown' else sc_from_road.get(sc, 'Unknown')
        to_map = {}
        for sc in scs:
            to = find_to(ws, hr, sc)
            to_map[sc] = to if to != 'Unknown' else sc_to_road.get(sc, 'Unknown')

        # Build ordered list of from_road groups preserving left-to-right order
        seen_fr = []
        groups  = {}
        for sc in scs:
            fr = fr_map[sc]
            if fr not in groups:
                groups[fr] = {'scs': [], 'to_roads': set()}
                seen_fr.append(fr)
            groups[fr]['scs'].append(sc)
            groups[fr]['to_roads'].add(to_map[sc])

        write_col = scs[0]  # start from leftmost original column
        row_groups[hr] = []

        for fr in seen_fr:
            g_scs     = groups[fr]['scs']
            to_roads  = sorted(groups[fr]['to_roads'])
            row_groups[hr].append((fr, g_scs, to_roads))

            for sc in g_scs:
                new_start_col[(hr, sc)] = write_col
                write_col += BLOCK_STRIDE

            summary_start[(hr, fr)] = write_col + SUMMARY_GAP
            write_col = summary_start[(hr, fr)] + _summary_width(to_roads) + SUMMARY_TRAIL

            # Reserve space for the adjacent To Road table (to: fr, all from-roads going there)
            all_row_scs = sorted(row_blocks[hr])
            to_road_from_roads = {fr_map[sc] for sc in all_row_scs if to_map[sc] == fr and fr_map[sc] != fr}
            if to_road_from_roads:
                write_col += _summary_width(list(to_road_from_roads)) + SUMMARY_TRAIL
                write_col += _summary_width([]) + SUMMARY_TRAIL  # Both table: Time + TOTAL + MaxHourly + Sum

    return new_start_col, summary_start, row_groups


new_start_col, summary_start, row_groups = compute_layout(block_positions, ws)

# Build orig->new mapping (consistent across rows for the same orig_sc)
orig_to_new_sc = {}
for (hr, orig_sc), nsc in new_start_col.items():
    orig_to_new_sc[orig_sc] = nsc
new_to_orig_sc = {v: k for k, v in orig_to_new_sc.items()}

  # Read raw cell data for any block that needs to move, then write + clear.

def _block_meta_start(ws, header_row, start_col):
    """Scan upward from header_row to find the topmost metadata row of this block.
    Does NOT stop at blank rows — metadata may be separated from the header by a blank."""
    meta_start = header_row
    for r in range(header_row - 1, max(header_row - 10, 0), -1):
        for c in range(start_col, start_col + BLOCK_WIDTH):
            v = normalize(ws.cell(r, c).value)
            if v in ('direction:', 'date:', 'station:', 'from:', 'to'):
                meta_start = r
                break
    return meta_start

def _read_block(ws, header_row, start_col):
    """Read all cells of a block (metadata through last data row) into memory.
    For metadata rows (above the header), formulas are resolved to literal values
    so the moved cells don't depend on cached formula results."""
    meta_start = _block_meta_start(ws, header_row, start_col)
    drows      = data_rows_of(ws, header_row, start_col)
    last_row   = drows[-1] if drows else header_row
    cells = []
    for r in range(meta_start, last_row + 1):
        is_meta = r < header_row
        for col_offset in range(len(HEADER)):
            cell = ws.cell(r, start_col + col_offset)
            if not isinstance(cell, MergedCell):
                value = cell.value
                if is_meta and isinstance(value, str) and value.startswith('='):
                    value = _resolve_cell(ws, r, start_col + col_offset)
                cells.append((r, col_offset, value, cell.number_format))
    return cells

def _write_block(ws, cells, orig_sc, new_sc):
    delta = new_sc - orig_sc
    for r, col_offset, value, fmt in cells:
        new_cell = ws.cell(r, new_sc + col_offset)
        if isinstance(new_cell, MergedCell):
            continue
        if isinstance(value, str) and value.startswith('=') and delta != 0:
            value = shift_formula_cols(value, delta)
        new_cell.value = value
        if fmt:
            new_cell.number_format = fmt

def _clear_block(ws, cells, orig_sc):
    rows_seen = set()
    for r, *_ in cells:
        rows_seen.add(r)
        for c_off in range(BLOCK_WIDTH):
            cell = ws.cell(r, orig_sc + c_off)
            if isinstance(cell, MergedCell):
                continue
            cell.value         = None
            cell.number_format = 'General'
            cell.fill          = _no_fill
            cell.border        = _no_border
            cell.font          = _no_font
    if rows_seen:
        last_data_row = max(rows_seen)
        for extra_r in (last_data_row + 1, last_data_row + 2):
            for c_off in range(BLOCK_WIDTH):
                cell = ws.cell(extra_r, orig_sc + c_off)
                if isinstance(cell, MergedCell):
                    continue
                cell.value         = None
                cell.number_format = 'General'
                cell.fill          = _no_fill
                cell.border        = _no_border
                cell.font          = _no_font

def _style_block(ws, header_row, start_col):
    from_to_labels = {'from:', 'from', 'to:', 'to'}
    meta_start = _block_meta_start(ws, header_row, start_col)
    drows = data_rows_of(ws, header_row, start_col)
    last_row = drows[-1] if drows else header_row
    for r in range(meta_start, last_row + 3):
        is_from_to = any(
            normalize(ws.cell(r, start_col + k).value) in from_to_labels
            for k in range(min(5, BLOCK_WIDTH))
        )
        is_data   = r > header_row
        if is_from_to:
            font = garamond_blue
        elif is_data:
            font = garamond_regular
        else:
            font = garamond_bold
        for col_offset in range(BLOCK_WIDTH):
            # skip computed cols — styled separately with red font
            if col_offset >= len(HEADER):
                continue
            col = start_col + col_offset
            cell = ws.cell(r, col)
            if isinstance(cell, MergedCell):
                continue
            if cell.value is None:
                continue
            safe_style(ws, r, col, font=font, alignment=wrap_center, border=thin_border)

# Collect moves; sort by new_sc descending so rightmost writes first
moves = [
    (hr, orig_sc, nsc)
    for (hr, orig_sc), nsc in new_start_col.items()
    if nsc != orig_sc
]
moves.sort(key=lambda x: -x[2])

# Pre-read all block data, then clear old positions, then write to new positions.
# This order avoids: (a) read-after-overwrite, (b) clear-after-write overlap.
move_data = {}
for hr, orig_sc, nsc in moves:
    move_data[(hr, orig_sc)] = _read_block(ws, hr, orig_sc)

for hr, orig_sc, nsc in moves:
    _clear_block(ws, move_data[(hr, orig_sc)], orig_sc)

for hr, orig_sc, nsc in moves:
    _write_block(ws, move_data[(hr, orig_sc)], orig_sc, nsc)

# Resolve metadata formula cells to literals for ALL blocks (moved and non-moved).
# Non-moved blocks retain formula strings from the source; resolve them so Excel
# shows the actual road name instead of a cell reference.
for hr, sc in block_positions:
    nsc = new_start_col.get((hr, sc), sc)
    meta_start = _block_meta_start(ws, hr, nsc)
    for r in range(meta_start, hr):
        for col_offset in range(BLOCK_WIDTH):
            cell = ws.cell(r, nsc + col_offset)
            if isinstance(cell, MergedCell):
                continue
            if isinstance(cell.value, str) and cell.value.startswith('='):
                cell.value = _resolve_cell(ws, r, nsc + col_offset)

# Resolve Hour-column time formulas (=X7+TIME(0,15,0) chains) to literal time values.
# _resolve_cell can't evaluate arithmetic; reconstruct times sequentially instead.
from datetime import timedelta
_15MIN = timedelta(minutes=15)
for hr, sc in block_positions:
    nsc  = new_start_col.get((hr, sc), sc)
    drows = data_rows_of(ws, hr, nsc)
    cur_t = None
    for r in drows:
        cell = ws.cell(r, nsc)
        v    = cell.value
        pt   = parse_time(v) if not (isinstance(v, str) and v.startswith('=')) else None
        if pt is not None:
            cur_t = pt
        elif cur_t is not None:
            from datetime import datetime as _dt
            cur_t = (_dt.combine(_dt.today(), cur_t) + _15MIN).time()
            cell.value = cur_t  # type: ignore
            cell.number_format = "h:mm AM/PM"

# Set adequate column widths for all moved-block columns so dates don't show ########.
for hr, sc in block_positions:
    nsc = new_start_col.get((hr, sc), sc)
    if nsc == sc:
        continue  # not moved; original width preserved
    for col_offset in range(BLOCK_WIDTH):
        col_letter = get_column_letter(nsc + col_offset)
        current = ws.column_dimensions[col_letter].width
        if current < 12:
            ws.column_dimensions[col_letter].width = 12

# ── Main block processing loop ─────────────────────────────────────────────
# Re-scan headers after relocation so start_col reflects new positions.

green_fill   = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
blue_fill    = PatternFill(start_color="4A90E2", end_color="4A90E2", fill_type="solid")
red_font      = Font(name='Garamond', bold=True, size=8, color='FF0000')
center_align  = Alignment(horizontal='center', vertical='center', wrap_text=True)
wrap_center   = center_align
garamond_bold    = Font(name='Garamond', bold=True, size=8)
garamond_regular = Font(name='Garamond', size=8)
garamond_red     = Font(name='Garamond', size=8, color='FF0000')
garamond_red_hdr = Font(name='Garamond', bold=True, size=8, color='FF0000')
garamond_blue    = Font(name='Garamond', bold=True, size=8, color='0000FF', underline='single')
yellow_fill  = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
header_fill  = PatternFill(start_color="4A90E2", end_color="4A90E2", fill_type="solid")
white_bold   = Font(name='Garamond', bold=True, size=8, color='FFFFFF')
thin_side    = Side(style="thin", color="808080")
medium_side  = Side(style="medium", color="000000")
thin_border  = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
heavy_border = Border(left=medium_side, right=medium_side, top=medium_side, bottom=medium_side)

block_totals           = {}
block_info_list        = []
col_period_peak_rows   = {}  # sc -> {pname: [cell_ref, ...]}
col_period_hourly_rows = {}  # sc -> {pname: [cell_ref, ...]}
block_registry      = []
col_layout          = {}
max_rows_per_col    = {}
global_last_max_row = 0

block_positions = find_headers(ws)  # re-scan after moves

for block_index, (header_row, start_col) in enumerate(block_positions):
    total_col  = start_col + len(HEADER)
    peak_col   = total_col + 1
    hourly_col = total_col + 2

    for i, title in enumerate(HEADER):
        ws.cell(header_row, start_col + i, value=title)
    ws.cell(header_row, total_col,  value="TOTAL")
    safe_style(ws, header_row, total_col,  font=garamond_red_hdr, alignment=wrap_center, border=thin_border)
    ws.cell(header_row, peak_col,   value="Peak Hourly Flow")
    safe_style(ws, header_row, peak_col,   font=garamond_red_hdr, alignment=wrap_center, border=thin_border)
    ws.cell(header_row, hourly_col, value="Hourly Flow")
    safe_style(ws, header_row, hourly_col, font=garamond_red_hdr, alignment=wrap_center, border=thin_border)

    vcols = [start_col + i for i in range(1, len(HEADER))]
    drows = data_rows_of(ws, header_row, start_col)
    rtot  = row_totals_of(ws, drows, vcols)

    first_v = get_column_letter(vcols[0])
    last_v  = get_column_letter(vcols[-1])
    for r in drows:
        cell = ws.cell(r, total_col)
        cell.value         = f"=SUM({first_v}{r}:{last_v}{r})"
        _formula_cache[cell.coordinate] = rtot[r]
        cell.number_format = 'General'
        safe_style(ws, r, total_col, font=garamond_red, alignment=wrap_center, border=thin_border)

    if not drows:
        continue

    total_row = drows[-1] + 1
    max_row   = total_row + 1

    if start_col not in max_rows_per_col:
        max_rows_per_col[start_col] = []
    max_rows_per_col[start_col].append(max_row)

    global_last_max_row = max(global_last_max_row, max_row)

    if start_col not in col_layout:
        col_layout[start_col] = {
            'total_col': total_col, 'peak_col': peak_col, 'hourly_col': hourly_col
        }

    safe_set(ws, total_row, start_col, "TOTAL")
    safe_set(ws, max_row,   start_col, "MAX")

    for col in vcols:
        col_l = get_column_letter(col)
        vals  = [to_num(ws.cell(r, col).value) for r in drows]
        s, m  = sum(vals), max(vals, default=0)
        safe_set(ws, total_row, col, f"=SUM({col_l}{drows[0]}:{col_l}{drows[-1]})")
        _set_cached(ws, total_row, col, int(s) if s == int(s) else s)
        safe_set(ws, max_row,   col, f"=MAX({col_l}{drows[0]}:{col_l}{drows[-1]})")
        _set_cached(ws, max_row, col, int(m) if m == int(m) else m)

    tot_tot = sum(rtot[r] for r in drows)
    max_tot = max(rtot[r] for r in drows)
    tot_l   = get_column_letter(total_col)
    safe_set(ws, total_row, total_col, f"=SUM({tot_l}{drows[0]}:{tot_l}{drows[-1]})")
    _set_cached(ws, total_row, total_col, int(tot_tot) if tot_tot == int(tot_tot) else tot_tot)
    safe_style(ws, total_row, total_col, number_format='General')
    safe_set(ws, max_row,   total_col, f"=MAX({tot_l}{drows[0]}:{tot_l}{drows[-1]})")
    _set_cached(ws, max_row, total_col, int(max_tot) if max_tot == int(max_tot) else max_tot)
    safe_style(ws, max_row,   total_col, number_format='General')

    block_info_list.append({'start_col': start_col, 'total_col': total_col, 'drows': drows})

    if start_col not in block_totals:
        block_totals[start_col] = []
    row_data = {}
    for offset, col in enumerate(vcols, start=1):
        col_sum = sum(to_num(ws.cell(r, col).value) for r in drows)
        row_data[offset] = (total_row, col, col_sum)
    row_data[13] = (total_row, total_col, tot_tot)
    block_totals[start_col].append(row_data)

    try:
        bus_col    = get_column_letter(start_col + HEADER.index("Bus"))
        others_col = get_column_letter(start_col + HEADER.index("Others"))
        tot_col_l  = get_column_letter(total_col)
        formula    = (f"=IFERROR(SUM({bus_col}{total_row}:{others_col}{total_row})"
                      f"/{tot_col_l}{total_row},0)")
        safe_set(ws, max_row, hourly_col, formula)
        safe_style(ws, max_row, hourly_col, number_format='0%')
    except ValueError:
        pass

    peak_letter   = get_column_letter(peak_col)
    hourly_letter = get_column_letter(hourly_col)
    for b in four_row_blocks(ws, drows, start_col, rtot):
        r_start = b['first_row']
        r_end   = b['last_row']
        for row in range(r_start, r_end + 1):
            safe_style(ws, row, peak_col,   fill=green_fill, border=thin_border)
            safe_style(ws, row, hourly_col, fill=blue_fill,  border=thin_border)
        try:
            ws.merge_cells(f"{peak_letter}{r_start}:{peak_letter}{r_end}")
        except Exception:
            pass
        try:
            ws.merge_cells(f"{hourly_letter}{r_start}:{hourly_letter}{r_end}")
        except Exception:
            pass
        safe_set(ws, r_start, peak_col,   f"=MAX({tot_l}{r_start}:{tot_l}{r_end})*4")
        _set_cached(ws, r_start, peak_col, int(b['peak']) if b['peak'] == int(b['peak']) else b['peak'])
        safe_style(ws, r_start, peak_col,   fill=green_fill, alignment=wrap_center,
                   font=garamond_red, number_format='General', border=thin_border)
        safe_set(ws, r_start, hourly_col, f"=SUM({tot_l}{r_start}:{tot_l}{r_end})")
        _set_cached(ws, r_start, hourly_col, int(b['hourly']) if b['hourly'] == int(b['hourly']) else b['hourly'])
        safe_style(ws, r_start, hourly_col, fill=blue_fill,  alignment=wrap_center,
                   font=garamond_red, number_format='General', border=thin_border)
        # Track which peak/hourly cells fall in each named period
        b_time = b['time']
        if b_time is not None:
            if start_col not in col_period_peak_rows:
                col_period_peak_rows[start_col]   = {p: [] for p in PERIODS}
                col_period_hourly_rows[start_col] = {p: [] for p in PERIODS}
            for pname, (sh, sm, eh, em) in PERIODS.items():
                if in_period(b_time, sh, sm, eh, em):
                    col_period_peak_rows[start_col][pname].append(
                        f"{peak_letter}{r_start}")
                    col_period_hourly_rows[start_col][pname].append(
                        f"{hourly_letter}{r_start}")

    _style_block(ws, header_row, start_col)

    time_totals = {}
    cur_t = None
    for r in range(header_row + 1, ws.max_row + 1):
        cell_val = ws.cell(r, start_col).value
        norm_val = normalize(cell_val)
        if norm_val in ('total', 'max', 'hour'):
            break
        if norm_val in ('direction:', 'date:', 'station:', 'from:', 'to'):
            break
        pt = parse_time(cell_val)
        if pt is not None:
            cur_t = pt
        elif cur_t is not None:
            mins = cur_t.hour * 60 + cur_t.minute + 15
            cur_t = time(mins // 60 % 24, mins % 60)
        if cur_t is None:
            continue
        row_total = sum(to_num(ws.cell(r, c).value) for c in vcols)
        time_totals[cur_t] = time_totals.get(cur_t, 0) + row_total

    block_registry.append({
        'from_road':     find_direction(ws, header_row, start_col),
        'to_road':       find_to(ws, header_row, start_col),
        'start_col':     start_col,
        'total_col':     total_col,
        'header_row':    header_row,
        'rightmost_col': hourly_col,
        'time_totals':   time_totals,
    })

# ── Summary rows (Morning/Lunch/Evening Peak, 3-day Average) ──────────────
morning_row   = global_last_max_row + 1
lunch_row     = morning_row + 1
evening_row   = lunch_row + 1
avg_daily_row = evening_row + 1

for sc, layout in col_layout.items():
    tc  = layout['total_col']
    pkc = layout['peak_col']
    hrc = layout['hourly_col']
    # Map new start_col back to original for Stage 1 period data lookup
    orig_sc = new_to_orig_sc.get(sc, sc)
    cpk = col_period_peaks.get(orig_sc,    {p: [] for p in PERIODS})
    cph = col_period_hourlies.get(orig_sc, {p: [] for p in PERIODS})

    safe_set(ws, morning_row,   tc, "Morning Peak")
    safe_style(ws, morning_row,   tc, font=garamond_bold, alignment=wrap_center)
    safe_set(ws, lunch_row,     tc, "Lunch Peak")
    safe_style(ws, lunch_row,     tc, font=garamond_bold, alignment=wrap_center)
    safe_set(ws, evening_row,   tc, "Evening Peak")
    safe_style(ws, evening_row,   tc, font=garamond_bold, alignment=wrap_center)
    safe_set(ws, avg_daily_row, sc, "3-day Average")
    safe_style(ws, avg_daily_row, sc, font=garamond_bold, alignment=wrap_center)

    for pname, row_idx in (
        ('morning', morning_row),
        ('lunch',   lunch_row),
        ('evening', evening_row),
    ):
        peak_refs   = col_period_peak_rows.get(sc, {}).get(pname, [])
        hourly_refs = col_period_hourly_rows.get(sc, {}).get(pname, [])
        if peak_refs:
            safe_set(ws, row_idx, pkc, f"=MAX({','.join(peak_refs)})")
            if cpk[pname]:
                _set_cached(ws, row_idx, pkc, int(max(cpk[pname])))
            safe_style(ws, row_idx, pkc, font=garamond_red, alignment=wrap_center)
        if hourly_refs:
            safe_set(ws, row_idx, hrc, f"=MAX({','.join(hourly_refs)})")
            if cph[pname]:
                _set_cached(ws, row_idx, hrc, int(max(cph[pname])))
            safe_style(ws, row_idx, hrc, font=garamond_red, alignment=wrap_center)

    block_tots = block_totals.get(sc, [])
    if block_tots:
        for offset in range(1, 14):
            refs = [
                f"{get_column_letter(bt[offset][1])}{bt[offset][0]}"
                for bt in block_tots if offset in bt
            ]
            if refs:
                nums    = [bt[offset][2] for bt in block_tots if offset in bt]
                avg_val = round(sum(nums) / len(nums)) if nums else 0
                formula = f"=ROUND(AVERAGE({','.join(refs)}),0)"
                safe_set(ws, avg_daily_row, sc + offset, formula)
                _set_cached(ws, avg_daily_row, sc + offset, avg_val)
                safe_style(ws, avg_daily_row, sc + offset, font=garamond_regular, alignment=wrap_center)
    
    # ── 3-day average for max % (HV formula) ──────────────────────────────
    max_rows = max_rows_per_col.get(sc, [])
    if max_rows:
        # Create a formula that averages the max % values from all max_rows
        hourly_col_letter = get_column_letter(hrc)
        cell_refs = [f"{hourly_col_letter}{mr}" for mr in max_rows]
        avg_formula = f"=IFERROR(AVERAGE({','.join(cell_refs)}),0)"
        safe_set(ws, avg_daily_row, hrc, avg_formula)
        safe_style(ws, avg_daily_row, hrc, font=garamond_regular, alignment=wrap_center, number_format='0%')

# ── Per-(from-road, day) summary tables ───────────────────────────────────
time_slots = []
h, m = 6, 0
while (h < 21) or (h == 21 and m <= 45):
    time_slots.append(time(h, m))
    m += 15
    if m >= 60:
        m -= 60
        h += 1

# Pre-pass: build cell_time_map for ALL blocks (needed by both From-Road and To-Road tables)
for b in block_registry:
    sc    = b['start_col']
    vcols = [sc + i for i in range(1, len(HEADER))]
    time_map, cur_t, first_time_row = {}, None, None
    for r in range(b['header_row'] + 1, ws.max_row + 1):
        cell_val = ws.cell(r, sc).value
        norm_val = normalize(cell_val)
        if norm_val in ('total', 'max', 'hour'):
            break
        if norm_val in ('direction:', 'date:', 'station:', 'from:', 'to'):
            break
        if norm_val and norm_val.startswith('direc'):
            break
        pt = parse_time(cell_val)
        if pt is not None:
            cur_t = pt
        elif cur_t is not None:
            mins = cur_t.hour * 60 + cur_t.minute + 15
            cur_t = time(mins // 60 % 24, mins % 60)
        if cur_t is None:
            continue
        if cur_t == time_slots[0] and first_time_row is None:
            first_time_row = r
        b_total_col = b.get('total_col', b['start_col'] + len(HEADER))
        b_vcols     = [b['start_col'] + i for i in range(1, len(HEADER))]
        row_val     = sum(to_num(ws.cell(r, c).value) for c in b_vcols)
        time_map.setdefault(cur_t, []).append((r, b_total_col, row_val))
    b['cell_time_map']  = time_map
    b['first_time_row'] = first_time_row

# Index registry by (header_row, from_road)
reg_groups = defaultdict(list)
for blk in block_registry:
    reg_groups[(blk['header_row'], blk['from_road'])].append(blk)

last_table_bottom = {}  # to_start_col -> bottom row of last table in that column band
summary_regions   = []  # collect all written summary tables for final restyle pass

def _restyle_summary_regions(ws, regions):
    for reg in regions:
        hr    = reg['header_row']
        sc    = reg['start_col']
        n_dest = reg['num_dest']
        tc    = reg['total_col']
        pkc   = reg['peak_col']
        sumc  = reg['sum_col']
        n_rows = reg['data_rows']
        for col in range(sc, sumc + 1):
            cell = ws.cell(hr, col)
            if not isinstance(cell, MergedCell) and cell.value is not None:
                safe_style(ws, hr, col, font=white_bold, fill=header_fill,
                           alignment=wrap_center, border=thin_border)
        for offset in range(1, n_rows + 1):
            r = hr + offset
            safe_style(ws, r, sc,  font=garamond_regular, border=thin_border, alignment=wrap_center)
            for ci in range(n_dest):
                safe_style(ws, r, sc + 1 + ci, font=garamond_regular, border=thin_border, alignment=wrap_center)
            safe_style(ws, r, tc,  font=garamond_regular, fill=yellow_fill, border=thin_border, alignment=wrap_center)
            if (offset - 1) % 4 == 0:
                safe_style(ws, r, pkc,  font=garamond_red, fill=yellow_fill, border=heavy_border, alignment=wrap_center)
                safe_style(ws, r, sumc, font=garamond_regular, fill=header_fill, border=thin_border, alignment=wrap_center)

for hr in sorted({b['header_row'] for b in block_registry}):
    for fr, orig_scs, to_roads_layout in row_groups.get(hr, []):
        group_blocks = reg_groups.get((hr, fr), [])
        if not group_blocks:
            continue

        to_start_col = summary_start.get((hr, fr))
        if to_start_col is None:
            continue

        to_roads_day = sorted({b['to_road'] for b in group_blocks})

        # Determine header row for this summary table
        first_rows = [b['first_time_row'] for b in group_blocks if b.get('first_time_row')]
        if first_rows:
            cur_row = min(first_rows) - 1
        else:
            end_row = hr
            for b in group_blocks:
                drows_hr = data_rows_of(ws, hr, b['start_col'])
                if drows_hr:
                    end_row = max(end_row, max(drows_hr))
            cur_row = end_row + 2

        prev_bottom = last_table_bottom.get(to_start_col, 0)
        if cur_row <= prev_bottom:
            cur_row = prev_bottom + 2

        total_col_day = to_start_col + 1 + len(to_roads_day)
        peak_col_day  = total_col_day + 1
        sum_col_day   = peak_col_day + 1

        # Write header row
        ws.cell(cur_row, to_start_col, value="Time")
        safe_style(ws, cur_row, to_start_col, fill=header_fill, alignment=center_align,
                   font=white_bold, border=thin_border)
        for ci, to_road in enumerate(to_roads_day):
            col = to_start_col + 1 + ci
            ws.cell(cur_row, col, value=f"To: {to_road}")
            safe_style(ws, cur_row, col, fill=header_fill, alignment=center_align,
                       font=white_bold, border=thin_border)
        ws.cell(cur_row, total_col_day, value="TOTAL")
        safe_style(ws, cur_row, total_col_day, fill=header_fill, alignment=center_align,
                   font=white_bold, border=thin_border)
        ws.cell(cur_row, peak_col_day, value="Max Hourly")
        safe_style(ws, cur_row, peak_col_day, fill=header_fill, alignment=center_align,
                   font=white_bold, border=thin_border)
        ws.cell(cur_row, sum_col_day, value="Sum")
        safe_style(ws, cur_row, sum_col_day, fill=header_fill, alignment=center_align,
                   font=white_bold, border=thin_border)

        # Write time rows
        slot_totals_from = []
        for offset, t in enumerate(time_slots, start=1):
            r = cur_row + offset
            ws.cell(r, to_start_col, value=t)
            safe_style(ws, r, to_start_col, number_format="h:mm AM/PM",
                       font=garamond_regular, border=thin_border, alignment=wrap_center)
            num_to_roads = len(to_roads_day)
            row_total    = 0
            for ci, to_road in enumerate(to_roads_day):
                col     = to_start_col + 1 + ci
                entries = [
                    (row, tc, val)
                    for b in group_blocks if b['to_road'] == to_road
                    for row, tc, val in b.get('cell_time_map', {}).get(t, [])
                ]
                refs        = [f"{get_column_letter(tc)}{row}" for row, tc, _ in entries]
                cell_num    = sum(val for _, _, val in entries)
                row_total  += cell_num
                cell        = ws.cell(r, col)
                cell.value  = (f"=SUM({','.join(refs)})" if len(refs) > 1
                               else f"={refs[0]}" if refs else 0)  # type: ignore
                if refs:
                    _formula_cache[cell.coordinate] = int(cell_num) if cell_num == int(cell_num) else cell_num
                cell.number_format = 'General'
                safe_style(ws, r, col, font=garamond_regular, border=thin_border, alignment=wrap_center)
            first_data_l = get_column_letter(to_start_col + 1)
            last_data_l  = get_column_letter(to_start_col + num_to_roads)
            total_cell             = ws.cell(r, total_col_day)
            total_cell.value       = f"=SUM({first_data_l}{r}:{last_data_l}{r})"  # type: ignore
            _formula_cache[total_cell.coordinate] = int(row_total) if row_total == int(row_total) else row_total
            total_cell.number_format = 'General'
            safe_style(ws, r, total_col_day, fill=yellow_fill, border=thin_border,
                       alignment=wrap_center, font=garamond_regular)
            slot_totals_from.append(row_total)

        # Write Peak Hourly and Sum (merged, per 4-row block)
        tot_day_l = get_column_letter(total_col_day)
        num_slots = len(time_slots)
        for i in range(0, num_slots, 4):
            chunk_size = min(4, num_slots - i)
            if not chunk_size:
                continue
            first_r    = cur_row + 1 + i
            last_r     = first_r + chunk_size - 1
            chunk      = slot_totals_from[i:i + chunk_size]
            peak_cache = int(max(chunk) * 4) if chunk else 0
            sum_cache  = int(sum(chunk)) if chunk else 0
            ws.merge_cells(start_row=first_r, start_column=peak_col_day,
                           end_row=last_r,   end_column=peak_col_day)
            safe_set(ws, first_r, peak_col_day,
                     f"=MAX({tot_day_l}{first_r}:{tot_day_l}{last_r})*4")
            _set_cached(ws, first_r, peak_col_day, peak_cache)
            safe_style(ws, first_r, peak_col_day, font=garamond_red, alignment=center_align,
                       fill=yellow_fill, border=heavy_border)
            ws.merge_cells(start_row=first_r, start_column=sum_col_day,
                           end_row=last_r,   end_column=sum_col_day)
            safe_set(ws, first_r, sum_col_day,
                     f"=SUM({tot_day_l}{first_r}:{tot_day_l}{last_r})")
            _set_cached(ws, first_r, sum_col_day, sum_cache)
            safe_style(ws, first_r, sum_col_day, alignment=wrap_center, fill=header_fill,
                       border=thin_border, font=garamond_regular)

        last_table_bottom[to_start_col] = cur_row + len(time_slots) + 1
        summary_regions.append({
            'header_row': cur_row,   'start_col': to_start_col,
            'num_dest':   len(to_roads_day),
            'total_col':  total_col_day, 'peak_col': peak_col_day,
            'sum_col':    sum_col_day,   'data_rows': len(time_slots),
        })
        print(f"  Summary: from='{fr}', day row={hr}, col={get_column_letter(to_start_col)}")

        # ── To Road summary table (placed immediately right of From Road summary) ──
        target_to_road    = fr
        tr_start_col      = sum_col_day + SUMMARY_TRAIL + 1
        to_road_blocks    = [
            b for b in block_registry
            if b['header_row'] == hr and b['to_road'] == target_to_road
        ]
        if to_road_blocks:
            from_roads_in = sorted(
                {b['from_road'] for b in to_road_blocks},
                key=lambda f: min(b['start_col'] for b in to_road_blocks if b['from_road'] == f)
            )
            tr_total_col = tr_start_col + 1 + len(from_roads_in)
            tr_peak_col  = tr_total_col + 1
            tr_sum_col   = tr_peak_col  + 1
            tr_cur_row   = cur_row

            ws.cell(tr_cur_row, tr_start_col, value=f"To: {target_to_road}")
            safe_style(ws, tr_cur_row, tr_start_col, fill=header_fill, alignment=center_align,
                       font=white_bold, border=thin_border)
            for ci, fr_name in enumerate(from_roads_in):
                col = tr_start_col + 1 + ci
                ws.cell(tr_cur_row, col, value=f"From: {fr_name}")
                safe_style(ws, tr_cur_row, col, fill=header_fill, alignment=center_align,
                           font=white_bold, border=thin_border)
            ws.cell(tr_cur_row, tr_total_col, value="TOTAL")
            safe_style(ws, tr_cur_row, tr_total_col, fill=header_fill, alignment=center_align,
                       font=white_bold, border=thin_border)
            ws.cell(tr_cur_row, tr_peak_col, value="Max Hourly")
            safe_style(ws, tr_cur_row, tr_peak_col, fill=header_fill, alignment=center_align,
                       font=white_bold, border=thin_border)
            ws.cell(tr_cur_row, tr_sum_col, value="Sum")
            safe_style(ws, tr_cur_row, tr_sum_col, fill=header_fill, alignment=center_align,
                       font=white_bold, border=thin_border)

            slot_totals_to = []
            for offset, t in enumerate(time_slots, start=1):
                r = tr_cur_row + offset
                ws.cell(r, tr_start_col, value=t)
                safe_style(ws, r, tr_start_col, number_format="h:mm AM/PM",
                           font=garamond_regular, border=thin_border, alignment=wrap_center)
                num_from_roads = len(from_roads_in)
                tr_row_total   = 0
                for ci, fr_name in enumerate(from_roads_in):
                    col     = tr_start_col + 1 + ci
                    entries = [
                        (row, tc, val)
                        for b in to_road_blocks if b['from_road'] == fr_name
                        for row, tc, val in b.get('cell_time_map', {}).get(t, [])
                    ]
                    refs        = [f"{get_column_letter(tc)}{row}" for row, tc, _ in entries]
                    cell_num    = sum(val for _, _, val in entries)
                    tr_row_total += cell_num
                    cell_t      = ws.cell(r, col)
                    cell_t.value = (f"=SUM({','.join(refs)})" if len(refs) > 1
                                    else f"={refs[0]}" if refs else 0)  # type: ignore
                    if refs:
                        _formula_cache[cell_t.coordinate] = int(cell_num) if cell_num == int(cell_num) else cell_num
                    cell_t.number_format = 'General'
                    safe_style(ws, r, col, font=garamond_regular, border=thin_border, alignment=wrap_center)
                tr_first_data_l = get_column_letter(tr_start_col + 1)
                tr_last_data_l  = get_column_letter(tr_start_col + num_from_roads)
                total_cell_t        = ws.cell(r, tr_total_col)
                total_cell_t.value  = f"=SUM({tr_first_data_l}{r}:{tr_last_data_l}{r})"  # type: ignore
                _formula_cache[total_cell_t.coordinate] = int(tr_row_total) if tr_row_total == int(tr_row_total) else tr_row_total
                total_cell_t.number_format = 'General'
                safe_style(ws, r, tr_total_col, fill=yellow_fill, border=thin_border,
                           alignment=wrap_center, font=garamond_regular)
                slot_totals_to.append(tr_row_total)

            tr_tot_l     = get_column_letter(tr_total_col)
            num_slots_tr = len(time_slots)
            for i in range(0, num_slots_tr, 4):
                chunk_size = min(4, num_slots_tr - i)
                if not chunk_size:
                    continue
                first_r     = tr_cur_row + 1 + i
                last_r      = first_r + chunk_size - 1
                tr_chunk    = slot_totals_to[i:i + chunk_size]
                tr_peak_c   = int(max(tr_chunk) * 4) if tr_chunk else 0
                tr_sum_c    = int(sum(tr_chunk)) if tr_chunk else 0
                ws.merge_cells(start_row=first_r, start_column=tr_peak_col,
                               end_row=last_r,   end_column=tr_peak_col)
                safe_set(ws, first_r, tr_peak_col,
                         f"=MAX({tr_tot_l}{first_r}:{tr_tot_l}{last_r})*4")
                _set_cached(ws, first_r, tr_peak_col, tr_peak_c)
                safe_style(ws, first_r, tr_peak_col, font=garamond_red, alignment=center_align,
                           fill=yellow_fill, border=heavy_border)
                ws.merge_cells(start_row=first_r, start_column=tr_sum_col,
                               end_row=last_r,   end_column=tr_sum_col)
                safe_set(ws, first_r, tr_sum_col,
                         f"=SUM({tr_tot_l}{first_r}:{tr_tot_l}{last_r})")
                _set_cached(ws, first_r, tr_sum_col, tr_sum_c)
                safe_style(ws, first_r, tr_sum_col, alignment=wrap_center, fill=header_fill,
                           border=thin_border, font=garamond_regular)

            summary_regions.append({
                'header_row': tr_cur_row,  'start_col': tr_start_col,
                'num_dest':   len(from_roads_in),
                'total_col':  tr_total_col, 'peak_col': tr_peak_col,
                'sum_col':    tr_sum_col,   'data_rows': len(time_slots),
            })
            print(f"  To-Road Summary: to='{target_to_road}', day row={hr}, col={get_column_letter(tr_start_col)}")

            # ── Both table (From + To combined) ──
            both_start_col = tr_sum_col + SUMMARY_TRAIL + 1
            both_total_col = both_start_col + 1
            both_peak_col  = both_total_col + 1
            both_sum_col   = both_peak_col  + 1

            ws.cell(tr_cur_row, both_start_col, value="Both")
            safe_style(ws, tr_cur_row, both_start_col, fill=header_fill, alignment=center_align,
                       font=white_bold, border=thin_border)
            ws.cell(tr_cur_row, both_total_col, value="TOTAL")
            safe_style(ws, tr_cur_row, both_total_col, fill=header_fill, alignment=center_align,
                       font=white_bold, border=thin_border)
            ws.cell(tr_cur_row, both_peak_col, value="Max Hourly")
            safe_style(ws, tr_cur_row, both_peak_col, fill=header_fill, alignment=center_align,
                       font=white_bold, border=thin_border)
            ws.cell(tr_cur_row, both_sum_col, value="Sum")
            safe_style(ws, tr_cur_row, both_sum_col, fill=header_fill, alignment=center_align,
                       font=white_bold, border=thin_border)

            slot_totals_both = []
            for offset, t in enumerate(time_slots, start=1):
                r = tr_cur_row + offset
                ws.cell(r, both_start_col, value=t)
                safe_style(ws, r, both_start_col, number_format="h:mm AM/PM",
                           font=garamond_regular, border=thin_border, alignment=wrap_center)
                from_ref   = f"{get_column_letter(total_col_day)}{r}"
                tr_ref     = f"{get_column_letter(tr_total_col)}{r}"
                both_val   = slot_totals_from[offset - 1] + slot_totals_to[offset - 1]
                both_cell  = ws.cell(r, both_total_col)
                both_cell.value = f"=SUM({from_ref},{tr_ref})"  # type: ignore
                _formula_cache[both_cell.coordinate] = int(both_val) if both_val == int(both_val) else both_val
                both_cell.number_format = 'General'
                safe_style(ws, r, both_total_col, fill=yellow_fill, border=thin_border,
                           alignment=wrap_center, font=garamond_regular)
                slot_totals_both.append(both_val)

            both_tot_l   = get_column_letter(both_total_col)
            num_slots_bt = len(time_slots)
            for i in range(0, num_slots_bt, 4):
                chunk_size = min(4, num_slots_bt - i)
                if not chunk_size:
                    continue
                first_r    = tr_cur_row + 1 + i
                last_r     = first_r + chunk_size - 1
                bt_chunk   = slot_totals_both[i:i + chunk_size]
                bt_peak_c  = int(max(bt_chunk) * 4) if bt_chunk else 0
                bt_sum_c   = int(sum(bt_chunk)) if bt_chunk else 0
                ws.merge_cells(start_row=first_r, start_column=both_peak_col,
                               end_row=last_r,   end_column=both_peak_col)
                safe_set(ws, first_r, both_peak_col,
                         f"=MAX({both_tot_l}{first_r}:{both_tot_l}{last_r})*4")
                _set_cached(ws, first_r, both_peak_col, bt_peak_c)
                safe_style(ws, first_r, both_peak_col, font=garamond_red, alignment=center_align,
                           fill=yellow_fill, border=heavy_border)
                ws.merge_cells(start_row=first_r, start_column=both_sum_col,
                               end_row=last_r,   end_column=both_sum_col)
                safe_set(ws, first_r, both_sum_col,
                         f"=SUM({both_tot_l}{first_r}:{both_tot_l}{last_r})")
                _set_cached(ws, first_r, both_sum_col, bt_sum_c)
                safe_style(ws, first_r, both_sum_col, alignment=wrap_center, fill=header_fill,
                           border=thin_border, font=garamond_regular)

            summary_regions.append({
                'header_row': tr_cur_row, 'start_col': both_start_col,
                'num_dest':   0,
                'total_col':  both_total_col, 'peak_col': both_peak_col,
                'sum_col':    both_sum_col,   'data_rows': len(time_slots),
            })
            print(f"  Both Summary: day row={hr}, col={get_column_letter(both_start_col)}")

_restyle_summary_regions(ws, summary_regions)

for col_idx in range(1, ws.max_column + 1):
    ws.column_dimensions[get_column_letter(col_idx)].width = 10

wb.calculation.fullCalcOnLoad = True
wb.save(output_path)
print(f"\nStage 2 complete.\nOutput -> {output_path}")
print(f"Processed {len(block_positions)} block(s).")
