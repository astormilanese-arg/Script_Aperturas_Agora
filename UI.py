# UI.py
import subprocess
import ast
import dearpygui.dearpygui as dpg

CONFIG_PATH = "Dependencies/config.py"
MAIN_SCRIPT = "Main.py"

config_data = {}

def parse_config_py():
    global config_data
    config_data = {}
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        for line in f:
            if "=" in line:
                key, val = line.split("=", 1)
                key = key.strip()
                val = ast.literal_eval(val.strip())
                config_data[key] = val

def write_config_py():
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        for key, val in config_data.items():
            val_str = repr(val)
            f.write(f"{key} = {val_str}\n")

def on_start_pressed():
    for key in config_data:
        config_data[key] = dpg.get_value(f"input_{key}")
    write_config_py()
    subprocess.Popen(["python", MAIN_SCRIPT])

def build_ui():
    dpg.create_context()
    dpg.create_viewport(title='Config Editor', width=800, height=400)
    dpg.setup_dearpygui()

    with dpg.window(label="Config Editor", width=780, height=350):
        dpg.add_text("Edit config variables:")
        parse_config_py()
        for key, val in config_data.items():
            dpg.add_input_text(label=key, default_value=val, width=600, tag=f"input_{key}")
        dpg.add_spacing(count=2)
        dpg.add_button(label="Start", callback=on_start_pressed)

    dpg.show_viewport()
    dpg.start_dearpygui()
    dpg.destroy_context()

build_ui()