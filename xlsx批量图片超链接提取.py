import os
import zipfile
import re
import requests
import concurrent.futures
import time
import imghdr

def analysis_url(text):
    urls = re.findall(r'https:[^"]+\.jpg', text)

    # 如果目录不存在，则创建
    if not os.path.exists(download_dir):
        os.makedirs(download_dir)
        print("下载目录已创建:", download_dir)
    else:
        print("下载目录已存在:", download_dir)
    
    return urls

def download_image(url):
    try:
        imageData = requests.get(url=url).content

        # 使用os.path.basename()获取文件名
        file_name = os.path.basename(url)
        with open(os.path.join(download_dir, file_name), 'wb') as f:
            f.write(imageData)
            print(file_name + ': 下载完成')
        
    except Exception as e:
        print("下载", url, "时出现错误:", e)
        print("将 URL 加入重试列表")
        return url

def run(urls):
    retry_list = []  # 存储需要重试的 URL

    # 创建线程池
    with concurrent.futures.ThreadPoolExecutor(max_workers=max_threads) as executor:
        # 提交下载任务
        future_to_url = {executor.submit(download_image, url): url for url in urls}

        # 处理已完成的任务
        for future in concurrent.futures.as_completed(future_to_url):
            url = future_to_url[future]
            try:
                future.result()  # 获取结果，如果有异常会抛出异常
            except Exception as exc:
                print('下载失败:', url)
                retry_list.append(url)

    print("所有图片下载完成")

    # 重试失败的下载
    print("开始重试失败的下载...")
    while retry_list:
        print("剩余重试任务数:", len(retry_list))
        new_retry_list = []
        with concurrent.futures.ThreadPoolExecutor(max_workers=max_threads) as executor:
            future_to_url = {executor.submit(download_image, url): url for url in retry_list}
            for future in concurrent.futures.as_completed(future_to_url):
                url = future_to_url[future]
                try:
                    future.result()
                except Exception as exc:
                    print('下载失败:', url)
                    new_retry_list.append(url)
        retry_list = new_retry_list
        print("暂停三秒后继续尝试重试失败的下载...")
        time.sleep(3)

    print("重试下载完成")

    # 进行图片完整性检查
    image_complete_check()

def reconnnect(urls, dirs):
    filtered_urls = [url for url in urls if not any(directory in url for directory in dirs)]
    return set(filtered_urls)

def image_complete_check():
    print("开始检查图片完整性")
    list = os.listdir(download_dir)

    count = 0
    for img in list:
        path = download_dir + img
        type = imghdr.what(path)
        if type == None:
            print(f'当前图片已损坏：{img}  已删除')
            os.remove(path)
            count += 1
    print(f"已删除损坏的图片：{count} 张")
    if count >= 1:
        print("准备重新执行下载程序")
        execute()


def execute():
    urls = analysis_url(text)
    re_urls = reconnnect(urls, os.listdir(download_dir))

    print("总任务数："+ str(len(urls)) + "; 剩余任务数：" + str(len(re_urls)))
    print("开始下载... ...")
    run(re_urls)

if __name__ == '__main__':

    print("**************请注意路径中不能包含中文、空格**************")
    print("该程序原理为：将xlsx文件后缀名改为zip并解压，在zip文件中存在<drawing1.xml.rels>文件，此文件存储表格中所有图片内容")
    print()
    # 从命令行参数中获取下载目录和文件名
    xlsx_file = input("请输入 XLSX 文件路径：")

    # 如果用户没有输入下载目录，则使用默认目录
    download_dir = input("请输入下载图片保存路径（默认当前目录下out文件夹）：")
    if not download_dir:
        download_dir = './out/'

    # 设置最大线程数
    max_threads = input("请设置最大线程数（默认10）：")
    if not max_threads:
        max_threads = 10
    else:
        max_threads = int(max_threads)

    # 解压缩XLSX文件为ZIP文件
    with zipfile.ZipFile(xlsx_file, "r") as zip_ref:
        # 获取ZIP文件中的文件列表
        files_in_zip = zip_ref.namelist()

        # 遍历文件列表查找drawing1.xml.rels文件
        for file_name in files_in_zip:
            if file_name == 'xl/drawings/_rels/drawing1.xml.rels':
                print("成功找到 drawing1.xml.rels 文件")
                print("开始解析")

                with zip_ref.open(file_name) as r:
                    text = r.read().decode('utf-8')
                execute()
                break
        else:
            print("未找到 drawing1.xml.rels 文件")
            print("退出程序")
