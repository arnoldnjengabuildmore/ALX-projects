"""Light-weight launcher to run the Tk GUI without a console (pythonw).

Double-click this file on Windows to run the GUI only.
"""
from pathlib import Path
import all_in_one_pipeline as app

script_root = Path(__file__).resolve().parent
input_dir, output_dir = app.ensure_io_folders(script_root)
default_input = app.prepare_default_input_file(script_root, input_dir)
default_output = output_dir / app.default_report_name_for_input(default_input)

app.launch_gui(default_input, default_output)
