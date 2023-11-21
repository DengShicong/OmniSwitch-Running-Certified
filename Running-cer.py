import os
import pandas as pd
from datetime import datetime
from netmiko import ConnectHandler
import re
import multiprocessing

def worker(device_info, output_queue):
    try:
        with ConnectHandler(**device_info) as net_connect:
            net_connect.enable()
            output = net_connect.send_command('show running-directory')
            running_config_lines = re.findall(r'Running configuration.*(?:\n(?!\n).*)*', output)
            if running_config_lines:
                formatted_output = "\n".join(running_config_lines)
                formatted_output = re.sub(r"WORKING", r"<span style='color:green;'>WORKING</span>", formatted_output)
                formatted_output = re.sub(r"CERTIFIED", r"<span style='color:red;'>CERTIFIED</span>", formatted_output)
                output_queue.put(f"{device_info['host']}:\n{formatted_output}\n\n")
    except Exception as e:
        output_queue.put(f"Connection failed for {device_info['host']}: {e}")

def execute_commands_and_save_to_html(template_path):
    template_data = pd.ExcelFile(template_path)
    assets_data = template_data.parse('assets')

    running_dir = "running-directory"
    os.makedirs(running_dir, exist_ok=True)

    running_directory_outputs = "<html><body><pre>"

    output_queue = multiprocessing.Queue()
    processes = []

    for _, row in assets_data.iterrows():
        device_info = {
            "device_type": row["device_type"],
            "host": row["IP"],
            "username": row["username"],
            "password": row["password"],
            "port": int(row["port"]) if pd.notna(row["port"]) else 22,
            "secret": row["secret"]
        }
        p = multiprocessing.Process(target=worker, args=(device_info, output_queue))
        processes.append(p)
        p.start()

    for p in processes:
        p.join()

    while not output_queue.empty():
        running_directory_outputs += output_queue.get()

    running_directory_outputs += "</pre></body></html>"

    if running_directory_outputs:
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        filename = f"{timestamp}_running-directory.html"
        filepath = os.path.join(running_dir, filename)
        with open(filepath, "w") as file:
            file.write(running_directory_outputs)

if __name__ == '__main__':
    template_path = 'template.xlsx'  # 替换为您模板文件的实际路径
    execute_commands_and_save_to_html(template_path)
