import subprocess
import json
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import time

LOG_FILE = "process_log.txt"  # 日志文件名

def write_log(message):
    """将信息写入日志文件"""
    with open(LOG_FILE, "a", encoding="utf-8") as log:
        log.write(message + "\n")

def parse_duration_to_seconds(duration):
    """
    将时间字符串 (HH:MM:SS.sss) 转换为浮点数秒数。
    支持毫秒部分，并返回完整的浮点数值。
    示例：1:10:10.1110000 -> 4210.111
    """
    try:
        # 分割时间字符串
        parts = duration.split(":")
        if len(parts) != 3:
            raise ValueError("Invalid duration format")
        
        hours, minutes, seconds = parts
        # 将小时和分钟转换为整数，秒部分可能包含小数
        total_seconds = int(hours) * 3600 + int(minutes) * 60 + float(seconds)
        return round(total_seconds, 4)  # 保留四位小数
    except Exception as e:
        write_log(f"Error parsing duration '{duration}': {e}")
        return None

def get_tag_durations_and_seconds(mkv_file):
    """提取视频和音频的时长"""
    mkvmerge_path = os.path.join(os.getcwd(), "mkvmerge.exe")
    if not os.path.isfile(mkvmerge_path):
        return None, None, "Error: mkvmerge.exe not found in the current directory."

    try:
        result = subprocess.run(
            [mkvmerge_path, "-J", mkv_file],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            errors="replace"  # 处理文件名中的特殊字符
        )
        if result.returncode != 0:
            return None, None, f"Error running mkvmerge: {result.stderr}"

        mkv_info = json.loads(result.stdout)
        tracks = mkv_info.get("tracks", [])
        video_time, audio_time = None, None

        for track in tracks:
            track_type = track.get("type", "unknown")
            tag_duration = track.get("properties", {}).get("tag_duration", "unknown")
            if tag_duration != "unknown":
                duration_in_seconds = parse_duration_to_seconds(tag_duration)
                if track_type == "video":
                    video_time = duration_in_seconds
                elif track_type == "audio":
                    audio_time = duration_in_seconds

        return video_time, audio_time, None

    except Exception as e:
        return None, None, f"Error processing mkvmerge output: {e}"

def sync_audio_video(input_file):
    """根据规则运行同步处理"""
    video_time, audio_time, error = get_tag_durations_and_seconds(input_file)
    if error:
        write_log(f"Error processing {input_file}: {error}")
        return f"Error processing {input_file}: {error}"

    log_message = f"File: {input_file}\nVideo Time: {video_time} s\nAudio Time: {audio_time} s\n"
    if video_time is not None and audio_time is not None:
        if video_time < audio_time:
            file_name, file_ext = os.path.splitext(input_file)
            output_file = f"{file_name}_1{file_ext}"
            mkvmerge_path = os.path.join(os.getcwd(), "mkvmerge.exe")
            sync_command = [
                mkvmerge_path,
                "-o", output_file,
                "--sync", f"1:0,{video_time}/{audio_time}",
                input_file
            ]
            try:
                result = subprocess.run(sync_command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
                if result.returncode == 0:
                    log_message += f"Rule: video_time < audio_time\nGenerated File: {output_file}\n"
                    write_log(log_message)
                    return f"Success: {input_file} -> {output_file}"
                else:
                    log_message += f"Error: {result.stderr}\n"
                    write_log(log_message)
                    return f"Error processing {input_file}: {result.stderr}"
            except Exception as e:
                log_message += f"Error: {e}\n"
                write_log(log_message)
                return f"Error executing mkvmerge for {input_file}: {e}"
        else:
            log_message += f"Rule: video_time == audio_time\nSkipped\n"
            write_log(log_message)
            return f"Skipped: {input_file} (video_time == audio_time)"
    log_message += "Error: Unable to retrieve durations\n"
    write_log(log_message)
    return f"Error: Unable to retrieve durations for {input_file}"

def process_folder(folder_path):
    """批量处理文件夹下的所有 .mkv 文件"""
    results = []
    for file in os.listdir(folder_path):
        if file.lower().endswith(".mkv"):
            file_path = os.path.join(folder_path, file)
            result = sync_audio_video(file_path)
            results.append(result)
    return results

def select_file():
    """选择单个文件并处理"""
    file_path = filedialog.askopenfilename(filetypes=[("MKV Files", "*.mkv")])
    if file_path:
        result = sync_audio_video(file_path)
        messagebox.showinfo("处理结果", result)

def select_folder():
    """选择文件夹并批量处理"""
    folder_path = filedialog.askdirectory()
    if folder_path:
        results = process_folder(folder_path)
        result_message = "\n".join(results)
        messagebox.showinfo("批量处理结果", result_message)

def main():
    """创建 GUI 界面"""
    root = tk.Tk()
    root.title("MKV 时间同步处理")
    root.geometry("400x200")

    frame = ttk.Frame(root, padding=10)
    frame.pack(fill="both", expand=True)

    label = ttk.Label(frame, text="选择操作：")
    label.pack(pady=10)

    single_button = ttk.Button(frame, text="选择单个文件", command=select_file)
    single_button.pack(pady=5)

    folder_button = ttk.Button(frame, text="选择文件夹", command=select_folder)
    folder_button.pack(pady=5)

    root.mainloop()

if __name__ == "__main__":
    # 初始化日志文件
    with open(LOG_FILE, "w", encoding="utf-8") as log:
        log.write(f"Log started at {time.strftime('%Y-%m-%d %H:%M:%S')}\n\n")
    main()
