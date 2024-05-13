"""
新版学习平台课程题目批量上传
EXCEL请遵循如下格式：
  题目  选项  答案  解析
  xx    x    x    x
"""

import requests
import pandas as pd


def analysis_excel() -> list[dict]:
    """
    解析excel题目文档并返回一个题目字典列表
    :return: list[dict...]
    """
    subject_list = []
    for i in range(len(excel)):
        subject = excel.loc[i, '题目']
        item = excel.loc[i, '选项'].replace('\n', "*").replace("**", '*')
        answer = excel.loc[i, '答案']
        explain = excel.loc[i, '解析'].replace('\n', '')

        data = {
            "id": "",
            "score": "2",  # 单题分数
            "difficulty": 1,  # 难度级别
            "answer": f"{answer}",  # 正确答案
            "typeInfo": 3,  # 所属范围
            "courseId": "15",  # 课程ID
            "item": f"{item}",  # 选项，每个选项之间用*分割
            "type": 2,  # 试卷类型 1：判断 2：单选 3：多选
            "subject": f"<p>{subject}</p>",  # 题目
            "explain": f"<p>{explain}</p>",  # 解析
            "remark": ""
        }

        subject_list.append(data)

    return subject_list


def batch_upload(subject_list: list, token: str) -> None:
    """
    批量上传题目
    :param subject_list: 待上传的题目列表
    :param token: 身份令牌
    :return: None
    """

    headers = {
        "Token": token
    }
    for i, sub in enumerate(subject_list):
        resp = requests.post(url, json=sub, headers=headers).json()
        if resp['code'] == 200:
            print(f"\033[92m {i}: 上传成功 \033[0m")
        else:
            print(f"\033[0;31;40m {i}: 上传失败 \033[0m")
        break


def login(username: str, password: str) -> str:
    """
    登录并获取token
    :param username: 账号
    :param password: 密码
    :return: str
    """
    login_url = "https://fangzhen.autotool.com.cn/api/LearnCourse/user/login"

    login_data = {
        "username": username,
        "password": password
    }

    login_info = requests.post(login_url, json=login_data).json()
    if login_info['code'] == 200:
        return login_info['data']['cookieId']


if __name__ == '__main__':
    url = "https://fangzhen.autotool.com.cn/api/LearnCourse/questions/insertOrUpdateQuestions"
    # excel文件路径
    excel_path = "C:\\Users\\K\\Desktop\\岗前培训题目.xlsx"
    excel = pd.read_excel(io=excel_path, usecols=['题目', '选项', '答案', '解析'])

    # 账号密码
    username = ""
    password = ""

    sub_list = analysis_excel()
    token = login(username, password)

    batch_upload(sub_list, token)
