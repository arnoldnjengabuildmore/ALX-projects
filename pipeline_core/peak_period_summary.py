"""
peak_period_summary.py

Creates a summary sheet with rows per direction and peak period using values
already written by `Final Calculation Flow.py` in the processed workbook.

Output columns:
1) Direction
2) Peak Period (Morning, Lunch, Evening)
3) Peak Hourly Flow
4) Hourly Flow
5) % Heavy Vehicles (from 3-day Average row)

Usage:
	python peak_period_summary.py --processed Raw Data 1_processed_YYYYMMDD_HHMMSS.xlsx

If `--processed` is not provided, the script picks the newest file containing
"_processed_" in the current working directory.
"""

from __future__ import annotations

import argparse
import os
from typing import Dict, List, Optional, Tuple

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


HEADER = [
	"Hour", "Motorcycles", "3-Wheelers/ Tuktuk", "Cars/Taxis/Jeeps/SUVs/Pick-ups",
	"Microbus", "Minibus", "Bus", "Omnibus", "LGV", "MGV", "HGV",
	"Articulated HGV", "Others",
]


def normalize(v) -> str:
	return "" if v is None else str(v).strip().lower()


def to_num(v) -> float:
	try:
		if v is None:
			return 0.0
		if isinstance(v, str):
			s = v.strip()
			if s == "":
				return 0.0
			return float(s)
		return float(v)
	except Exception:
		return 0.0


def find_headers(ws) -> List[Tuple[int, int]]:
	hits: List[Tuple[int, int]] = []
	width = len(HEADER)
	for r in range(1, ws.max_row + 1):
		for c in range(1, ws.max_column - width + 2):
			if all(normalize(ws.cell(r, c + i).value) == normalize(HEADER[i]) for i in range(width)):
				hits.append((r, c))
	hits.sort(key=lambda x: (x[0], x[1]))
	return hits


def find_direction(ws, header_row: int, start_col: int) -> str:
	meta = {"from:", "from", "to:", "to", "direction:", "date:", "station:", "hour", "total", "max"}
	from_val = "Unknown"
	to_val = "Unknown"
	for r in range(header_row - 1, max(header_row - 10, 0), -1):
		for c in range(start_col, start_col + 20):
			tag = normalize(ws.cell(r, c).value)
			if tag in ("from:", "from"):
				for cc in range(c + 1, c + 12):
					v = ws.cell(r, cc).value
					nv = normalize(v)
					if nv and nv not in meta:
						from_val = str(v).strip()
						break
			if tag in ("to:", "to"):
				for cc in range(c + 1, c + 12):
					v = ws.cell(r, cc).value
					nv = normalize(v)
					if nv and nv not in meta:
						to_val = str(v).strip()
						break
	if from_val == "Unknown" and to_val == "Unknown":
		return "Unknown"
	return f"{from_val} -> {to_val}"


def find_row_by_label(ws, col: int, label: str) -> Optional[int]:
	needle = normalize(label)
	for r in range(1, ws.max_row + 1):
		if normalize(ws.cell(r, col).value) == needle:
			return r
	return None


def find_row_by_label_in_range(ws, col: int, label: str, start_row: int, end_row: int) -> Optional[int]:
	needle = normalize(label)
	for r in range(max(1, start_row), min(ws.max_row, end_row) + 1):
		if normalize(ws.cell(r, col).value) == needle:
			return r
	return None


def latest_processed_file(cwd: str) -> Optional[str]:
	candidates = [
		f for f in os.listdir(cwd)
		if f.lower().endswith(".xlsx")
		and "_processed_" in f.lower()
		and not os.path.basename(f).startswith("~$")
	]
	if not candidates:
		return None
	candidates.sort(key=lambda name: os.path.getmtime(os.path.join(cwd, name)), reverse=True)
	return os.path.join(cwd, candidates[0])


def compute_hv_percent_like_avg_formula(ws, sc: int) -> float:
	"""Replicate Final Calculation Flow's avg_formula for %HV.

	Equivalent to averaging % values from all MAX rows in this block column.
	Each MAX-row % is computed as SUM(Bus:Others on TOTAL row) / TOTAL.
	"""
	max_rows: List[int] = []
	for r in range(1, ws.max_row + 1):
		if normalize(ws.cell(r, sc).value) == "max":
			max_rows.append(r)

	if not max_rows:
		return 0.0

	total_col = sc + len(HEADER)
	bus_offset = HEADER.index("Bus")
	others_offset = HEADER.index("Others")
	percentages: List[float] = []

	for max_row in max_rows:
		total_row = max_row - 1
		if total_row < 1:
			continue
		heavy_sum = 0.0
		for col in range(sc + bus_offset, sc + others_offset + 1):
			heavy_sum += to_num(ws.cell(total_row, col).value)
		total_val = to_num(ws.cell(total_row, total_col).value)
		if total_val > 0:
			percentages.append(heavy_sum / total_val)

	if not percentages:
		return 0.0
	return sum(percentages) / len(percentages)


def style_summary_sheet(ws, max_row: int) -> None:
	header_fill = PatternFill("solid", fgColor="D9EAF7")
	title_fill = PatternFill("solid", fgColor="F4F6FA")
	thin = Side(style="thin", color="7F7F7F")
	border = Border(left=thin, right=thin, top=thin, bottom=thin)

	ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=5)
	title = ws.cell(1, 1, "Peak Period Summary")
	title.fill = title_fill
	title.font = Font(bold=True, color="C00000", size=12)
	title.alignment = Alignment(horizontal="left", vertical="center")

	headers = ["Direction", "Peak Period", "Peak Hourly Flow", "Hourly Flow", "% Heavy Vehicles"]
	for col, text in enumerate(headers, start=1):
		c = ws.cell(2, col, text)
		c.fill = header_fill
		c.font = Font(bold=True, size=9)
		c.border = border
		c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

	for r in range(3, max_row + 1):
		for cidx in range(1, 6):
			c = ws.cell(r, cidx)
			c.border = border
			if cidx == 1:
				c.alignment = Alignment(horizontal="left", vertical="center")
			else:
				c.alignment = Alignment(horizontal="center", vertical="center")

	ws.column_dimensions["A"].width = 35
	ws.column_dimensions["B"].width = 14
	ws.column_dimensions["C"].width = 18
	ws.column_dimensions["D"].width = 14
	ws.column_dimensions["E"].width = 16
	ws.row_dimensions[1].height = 24
	ws.row_dimensions[2].height = 28


def main() -> None:
	parser = argparse.ArgumentParser()
	parser.add_argument("--processed", "-p", default=None, help="Processed workbook from Final Calculation Flow.py")
	parser.add_argument("--output", "-o", default=None, help="Output workbook path")
	args = parser.parse_args()

	cwd = os.getcwd()
	source_path = args.processed or latest_processed_file(cwd)
	if not source_path:
		raise FileNotFoundError("No processed workbook found. Pass --processed <file>.xlsx")
	if not os.path.exists(source_path):
		raise FileNotFoundError(f"Processed workbook not found: {source_path}")

	out_path = args.output or os.path.splitext(source_path)[0] + "_peak_period_summary.xlsx"

	wb = load_workbook(source_path, data_only=True)

	source_ws = None
	headers = []
	for ws in wb.worksheets:
		hits = find_headers(ws)
		if hits:
			source_ws = ws
			headers = hits
			break
	if source_ws is None:
		raise ValueError("No traffic header blocks found in workbook")

	period_order = ["Morning Peak", "Lunch Peak", "Evening Peak"]
	period_lookup = {
		"Morning Peak": "morning peak",
		"Lunch Peak": "lunch peak",
		"Evening Peak": "evening peak",
	}
	direction_rows: Dict[str, Dict[str, Tuple[float, float]]] = {}
	direction_hv: Dict[str, float] = {}
	direction_order: List[str] = []

	seen_keys = set()
	for header_row, sc in headers:
		direction = find_direction(source_ws, header_row, sc)
		# Deduplicate identical blocks that may appear in repeated layouts.
		dedupe_key = (header_row, sc, direction)
		if dedupe_key in seen_keys:
			continue
		seen_keys.add(dedupe_key)
		if direction not in direction_rows:
			direction_rows[direction] = {}
			direction_order.append(direction)

		total_col = sc + len(HEADER)
		peak_col = total_col + 1
		hourly_col = total_col + 2

		avg_row = find_row_by_label(source_ws, sc, "3-day average")
		hv_percent = compute_hv_percent_like_avg_formula(source_ws, sc)
		if direction not in direction_hv or direction_hv[direction] == 0:
			direction_hv[direction] = hv_percent

		# Prefer block-local rows relative to 3-day average:
		# morning=avg-3, lunch=avg-2, evening=avg-1 in Final Calculation Flow output.
		for label in period_order:
			lookup = period_lookup[label]
			period_row: Optional[int] = None
			if avg_row is not None:
				candidate = avg_row - {"Morning Peak": 3, "Lunch Peak": 2, "Evening Peak": 1}[label]
				if candidate >= 1 and normalize(source_ws.cell(candidate, total_col).value) == lookup:
					period_row = candidate
			if period_row is None:
				search_start = (avg_row - 12) if avg_row else 1
				search_end = (avg_row + 2) if avg_row else source_ws.max_row
				period_row = find_row_by_label_in_range(source_ws, total_col, lookup, search_start, search_end)
			if period_row is None:
				continue
			peak_hourly = to_num(source_ws.cell(period_row, peak_col).value)
			hourly_flow = to_num(source_ws.cell(period_row, hourly_col).value)
			if label not in direction_rows[direction]:
				direction_rows[direction][label] = (peak_hourly, hourly_flow)

	if "Peak Period Summary" in wb.sheetnames:
		del wb["Peak Period Summary"]
	out_ws = wb.create_sheet("Peak Period Summary")

	write_row = 3
	for direction in direction_order:
		start_row = write_row
		hv_percent = direction_hv.get(direction, 0.0)
		hv_percent = round(hv_percent * 100) / 100.0  # nearest 1%
		for period in period_order:
			if period not in direction_rows.get(direction, {}):
				continue
			peak_hourly, hourly_flow = direction_rows[direction][period]
			out_ws.cell(write_row, 1, direction)
			out_ws.cell(write_row, 2, period)
			out_ws.cell(write_row, 3, round(peak_hourly, 0))
			out_ws.cell(write_row, 4, round(hourly_flow, 0))
			hv_cell = out_ws.cell(write_row, 5, hv_percent)
			hv_cell.number_format = "0%"
			write_row += 1
		end_row = write_row - 1
		if end_row > start_row:
			out_ws.merge_cells(start_row=start_row, start_column=1, end_row=end_row, end_column=1)
			out_ws.merge_cells(start_row=start_row, start_column=5, end_row=end_row, end_column=5)

	style_summary_sheet(out_ws, max_row=max(write_row - 1, 2))
	wb.save(out_path)
	print(f"Peak Period Summary written to: {out_path}")


if __name__ == "__main__":
	main()
