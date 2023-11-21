import os
import pandas as pd
from datetime import datetime
from netmiko import ConnectHandler
import re

def execute_commands_and_save_to_html(template_path):
    template_data = pd.ExcelFile(template_path)
    assets_data = template_data.parse('assets')

    running_dir = "running-directory"
    os.makedirs(running_dir, exist_ok=True)

    running_directory_outputs = "<html><body><pre>"

    for index, row in assets_data.iterrows():
        device_info = {
            "device_type": row["device_type"],
            "host": row["IP"],
            "username": row["username"],
            "password": row["password"],
            "port": int(row["port"]) if pd.notna(row["port"]) else 22,
            "secret": row["secret"]
        }

        try:
            with ConnectHandler(**device_info) as net_connect:
                net_connect.enable()

                output = net_connect.send_command('show running-directory')

                running_config_lines = re.findall(r'Running configuration.*(?:\n(?!\n).*)*', output)
                if running_config_lines:
                    formatted_output = "\n".join(running_config_lines)
                    formatted_output = re.sub(r"WORKING", r"<span style='color:green;'>WORKING</span>", formatted_output)
                    formatted_output = re.sub(r"CERTIFIED", r"<span style='color:red;'>CERTIFIED</span>", formatted_output)
                    running_directory_outputs += f"{row['IP']}:\n{formatted_output}\n\n"

        except Exception as e:
            print(f"Connection failed for {row['IP']}: {e}")

    running_directory_outputs += "</pre></body></html>"

    if running_directory_outputs:
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        filename = f"{timestamp}_running-directory.html"
        filepath = os.path.join(running_dir, filename)
        with open(filepath, "w") as file:
            file.write(running_directory_outputs)

template_path = 'template.xlsx'  # 替换为您模板文件的实际路径
execute_commands_and_save_to_html(template_path)