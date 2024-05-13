import threading
from moviepy.editor import *
import os

print("*************************************\n"
      "            视频编辑 V0.1            \n"
      "     该工具目前仅用于批量删除视频音频     \n"
      "\n"
      "             注意事项：               \n"
      "     无需设置默认选项请直接回车确认       \n"
      "        线程数设置越大效率越高           \n"
      "        请按照电脑配置进行选择           \n"
      "           推荐默认10线程              \n"
      "*************************************\n")

dir_path = input("请输入原视频目录路径：").strip()
save_path = input("请输入新视频保存路径【默认当前盘符根目录】：").strip()
n = input("请设置线程数【默认线程：10】：")

if save_path == "":
    save_path = os.path.join(dir_path[:3], "/无声版01")

    if not os.path.exists(save_path):
        os.mkdir(save_path)

if n != "":
    n = int(n)

if n == "":
    n = 10


def remove_audio(video_list):
    for paths in video_list:
        video = VideoFileClip(paths[0])
        video = video.without_audio()
        video.write_videofile(paths[1], threads=18)


if __name__ == '__main__':
    dir = os.listdir(dir_path)

    task_list = []
    # 创建文件夹
    for root, dirs, names in os.walk(dir_path):
        for name in names:
            ext = os.path.splitext(name)[1]  # 获取后缀名
            if "MP4" in name:
                new_dir = (save_path + root.split(":")[1]).replace("\\", "/")
                os.makedirs(new_dir, exist_ok=True)
                task_list.append((os.path.join(root, name), os.path.join(new_dir, name)))

    # 拆分线程任务
    step = int(len(task_list) / n)
    task_list = [task_list[i:i + step] for i in range(0, len(task_list), step)]

    for task in task_list:
        threading.Thread(target=remove_audio, args=(task,)).start()
