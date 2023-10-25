# Script Name: LOL 汉化程序
# Author: 飘呀飘
# Version: 1.1.0
# Description: This script allows users to modify shortcuts for League of Legends clients to support Chinese localization.
# Date: 2023-10-24

import os
import pythoncom
import win32com.client
from win32com.client import Dispatch
import tkinter as tk
from tkinter import filedialog, messagebox
import psutil

version_number = "1.1.0"
desktop_dir = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
shortcut_path = os.path.join(desktop_dir, "LeagueClient - Shortcut.lnk")
user_home_dir = os.path.expanduser("~")
script_data_dir = os.path.join(user_home_dir, "LOLPathCN")
if not os.path.exists(script_data_dir):
    os.makedirs(script_data_dir)
last_file_path_file = os.path.join(script_data_dir, "last_file_path.txt")
last_pbe_file_path_file = os.path.join(script_data_dir, "last_pbe_file_path.txt")
last_account_pw = os.path.join(script_data_dir, "last_acc_pw.txt")
last_pbe_account_pw = os.path.join(script_data_dir, "last_pbe_acc_pw.txt")

def get_shortcut_target(shortcut_path):
    shell = win32com.client.Dispatch("WScript.Shell")
    shortcut = shell.CreateShortCut(shortcut_path)
    return shortcut.Targetpath


def close_riot_client_services():
    wmi = win32com.client.GetObject('winmgmts:')
    processes = wmi.InstancesOf('Win32_Process')
    for process in processes:
        if process.Properties_('Name').Value == "RiotClientServices.exe":
            pid = process.Properties_('ProcessId').Value
            try:
                psutil.Process(pid).terminate()
            except psutil.NoSuchProcess:
                pass
            except psutil.AccessDenied:
                pass
                break

def update_status_label():
    status_label_text = "当前: "
    if os.path.exists(shortcut_path):
        target_path = get_shortcut_target(shortcut_path)
        if "pbe" in target_path.lower():
            status_label_text += "PBE"
        else:
            status_label_text += "正式服"
    else:
        status_label_text += "未知"
    return status_label_text

def browse_file(entry_file_path):
    file_path = filedialog.askopenfilename()
    entry_file_path.delete(0, tk.END)
    entry_file_path.insert(0, file_path)
    save_last_file_path(file_path)

def browse_pbe_file(entry_pbe_file_path):
    file_path = filedialog.askopenfilename()
    entry_pbe_file_path.delete(0, tk.END)
    entry_pbe_file_path.insert(0, file_path)
    save_last_pbe_file_path(file_path)

def save_last_file_path(file_path):
    # Save the file path in the same directory as the script
    with open(last_file_path_file, 'w') as file:
        file.write(file_path)

def save_last_pbe_file_path(file_path):
    # Save the file path in the same directory as the script
    with open(last_pbe_file_path_file, 'w') as file:
        file.write(file_path)

def load_last_file_path():
    if os.path.exists(last_file_path_file):
        with open(last_file_path_file, 'r') as file:
            return file.readline().strip()
    return ""

def load_last_pbe_file_path():
    if os.path.exists(last_pbe_file_path_file):
        with open(last_pbe_file_path_file, 'r') as file:
            return file.readline().strip()
    return ""

def modify_shortcut_target(shortcut_path, target_path, arguments):
    # Create a shell object to manipulate the shortcut
    shell = Dispatch('WScript.Shell')
    shortcut = shell.CreateShortCut(shortcut_path)

    # Modify the target and arguments of the shortcut
    shortcut.Targetpath = target_path
    shortcut.Arguments = f'{arguments}'

    # Save the changes to the shortcut
    shortcut.save() 

def run_lol_shortcut():
    os.startfile(os.path.join(desktop_dir, "LeagueClient - Shortcut.lnk"))
        
def create_shortcut(target_file_path, shortcut_path, status_label):
    if "LeagueClient" not in target_file_path:
        error_msg = "Error: The selected file name should be 'LeagueClient'."
        messagebox.showerror("Error", error_msg)
        return False
    close_riot_client_services()
    if os.path.exists(shortcut_path):
        os.remove(shortcut_path)    
    try:
        pythoncom.CoInitialize()
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(shortcut_path)
        shortcut.Targetpath = target_file_path
        shortcut.save()
        modify_shortcut_target(shortcut_path, target_file_path, "--locale=zh_CN")
        status_label.config(text=update_status_label())
        run_lol_shortcut()
        return True
    except Exception as e:
        return False

def callback(url):
    import webbrowser
    webbrowser.open_new(url)

def create_main_window():
    # 创建主窗口
    root = tk.Tk()
    root.title("飘呀飘的汉化小程序")

    # 添加标签和输入框（LOL文件路径）
    label = tk.Label(root, text="Enter LOL File Path:")
    label.grid(row=0, column=0, pady=10)

    last_path = load_last_file_path()
    entry_file_path = tk.Entry(root, width=50, textvariable=tk.StringVar(value=last_path))
    entry_file_path.grid(row=1, column=0, pady=10)

    # 添加浏览按钮（LOL文件路径）
    browse_button = tk.Button(root, text="Browse", command=lambda: browse_file(entry_file_path))
    browse_button.grid(row=1, column=2, pady=10, padx=5)

    # 添加标签和输入框（PBE文件路径）
    label_pbe = tk.Label(root, text="Enter PBE File Path:")
    label_pbe.grid(row=2, column=0, pady=10)

    last_pbe_path = load_last_pbe_file_path()
    entry_pbe_file_path = tk.Entry(root, width=50, textvariable=tk.StringVar(value=last_pbe_path))
    entry_pbe_file_path.grid(row=3, column=0, pady=10)

    # 添加浏览按钮（PBE文件路径）
    browse_pbe_button = tk.Button(root, text="Browse", command=lambda: browse_pbe_file(entry_pbe_file_path))
    browse_pbe_button.grid(row=3, column=2, pady=10, padx=5)

    # 添加状态标签
    status_label = tk.Label(root, text="", font=("Arial", 10))
    status_label.grid(row=4, column=0, sticky="sw", padx=10, pady=10)

    # Call update_status_label to set the initial status label text
    status_label.config(text=update_status_label())

    # 添加LOL Run按钮
    run_lol_button = tk.Button(root, text="运行 正式服", command=lambda: create_shortcut(entry_file_path.get(), shortcut_path, status_label))
    run_lol_button.grid(row=4, column=0, columnspan=3, pady=10)

    # 添加PBE Run按钮
    run_pbe_button = tk.Button(root, text="运行 PBE", command=lambda: create_shortcut(entry_pbe_file_path.get(), shortcut_path, status_label))
    run_pbe_button.grid(row=5, column=0 , columnspan=3, pady=10)

    # 添加你的GitHub链接
    github_url = "https://github.com/hhxjqm"  # Replace with your actual GitHub URL
    github_link_label = tk.Label(root, text="GitHub", font=("Arial", 12), fg="blue", cursor="hand2")
    github_link_label.grid(row=5, column=2, sticky="se", padx=10, pady=10)
    github_link_label.bind("<Button-1>", lambda e: callback(github_url))

    # 添加版本号
    version_label = tk.Label(root, text=f"Version: {version_number}", font=("Arial", 10))
    version_label.grid(row=5, column=0, sticky="sw", padx=10, pady=10)

    root.mainloop()

def main():
    create_main_window()

if __name__ == "__main__":
    main()
