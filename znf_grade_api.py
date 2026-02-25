"""
按油猴脚本逻辑抓取正方“成绩详情导出”接口数据：
- POST application/x-www-form-urlencoded
- 依次尝试两条路径：/cjcx/... 与 /jwglxt/cjcx/...
- 不保存文件：将返回的 Excel 二进制解析为 pandas.DataFrame -> list[dict]
- 合并同一门课（同一教学班ID优先），并把“总评”放在成绩分项最前
"""

from __future__ import annotations

import os
import time
from io import BytesIO
from urllib.parse import urljoin
from collections import OrderedDict
from typing import Any

import requests
import pandas as pd

import requests
import base64
import json
import traceback
import binascii
import rsa
from urllib.parse import urljoin
from pyquery import PyQuery as pq

import sys
sys.stdout.reconfigure(encoding="utf-8")


# -----------------------------
# 1) 构造表单参数（完全按油猴脚本）
# -----------------------------
def build_form_body(xnm: str, xqm: str) -> list[tuple[str, str]]:
    """
    按油猴脚本逻辑构造 form-urlencoded 参数。
    使用 list[tuple] 以支持重复 key：exportModel.selectCol 多次出现。
    """
    params: list[tuple[str, str]] = []
    params.append(("gnmkdmKey", "N305005"))
    params.append(("xnm", xnm))
    params.append(("xqm", xqm))
    params.append(("dcclbh", "JW_N305005_GLY"))

    cols = [
        "xnmmc@学年",
        "xqmmc@学期",
        "jxb_id@教学班ID",
        "xf@学分",
        "kcmc@课程名称",
        "xmcj@成绩",
        "xmblmc@成绩分项",
    ]
    for col in cols:
        params.append(("exportModel.selectCol", col))

    params.append(("exportModel.exportWjgs", "xls"))
    params.append(("fileName", "成绩单"))
    return params

def encrypt_password(pwd, n, e):
    """教务系统专用RSA加密"""
    message = str(pwd).encode()

    # base64 → hex → int
    rsa_n = binascii.b2a_hex(binascii.a2b_base64(n))
    rsa_e = binascii.b2a_hex(binascii.a2b_base64(e))

    key = rsa.PublicKey(int(rsa_n, 16), int(rsa_e, 16))

    # RSA加密
    encropy_pwd = rsa.encrypt(message, key)

    # 返回base64
    result = binascii.b2a_base64(encropy_pwd).decode().strip()

    return result


def login_and_get_cookie(base_url, sid, password, timeout=5):
    """
    登录教务系统并返回cookies

    返回:
        dict cookies
    """

    session = requests.Session()

    headers = {
        "User-Agent": "Mozilla/5.0 Chrome/120.0 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    }

    login_url = urljoin(base_url, "xtgl/login_slogin.html")
    key_url = urljoin(base_url, "xtgl/login_getPublicKey.html")

    try:

        # Step 1 获取登录页
        resp = session.get(login_url, headers=headers, timeout=timeout)

        if resp.status_code != 200:
            print("登录页访问失败")
            return None

        doc = pq(resp.text)

        csrf_token = doc("#csrftoken").attr("value")

        if not csrf_token:
            print("获取csrf_token失败")
            return None

        # Step 2 获取publicKey
        pub = session.get(key_url, headers=headers, timeout=timeout).json()

        modulus = pub["modulus"]
        exponent = pub["exponent"]

        # Step 3 加密密码
        encrypt_pwd = encrypt_password(password, modulus, exponent)

        # Step 4 登录
        login_data = {
            "csrftoken": csrf_token,
            "yhm": sid,
            "mm": encrypt_pwd
        }

        headers["Referer"] = login_url

        login_resp = session.post(
            login_url,
            headers=headers,
            data=login_data,
            timeout=timeout
        )

        doc = pq(login_resp.text)
        tips = doc("p#tips").text()

        if "用户名或密码不正确" in tips:
            print("用户名或密码错误")
            return None

        # Step 5 获取cookies
        cookies = session.cookies.get_dict()

        # print("登录成功")

        return cookies

    except Exception as e:
        traceback.print_exc()
        return None

# -----------------------------
# 2) 发请求：尝试两条路径
# -----------------------------
def fetch_export_bytes(
    base_url: str,
    xnm: str,
    xqm: str,
    cookie: str,
    timeout: int = 30,
) -> bytes:
    """
    向正方导出接口发 POST，返回 Excel 二进制内容。
    注意：多数学校必须携带已登录 cookie。
    """
    session = requests.Session()

    headers = {
        "Content-Type": "application/x-www-form-urlencoded",
        "User-Agent": "Mozilla/5.0",
        "Cookie": cookie,
        "Referer": urljoin(base_url.rstrip("/") + "/", "cjcx/"),
    }

    body = build_form_body(xnm, xqm)

    paths = [
        "/cjcx/cjcx_dcXsKccjList.html",
        "/jwglxt/cjcx/cjcx_dcXsKccjList.html",
    ]

    last_err: Exception | None = None
    for path in paths:
        try:
            url = urljoin(base_url.rstrip("/") + "/", path.lstrip("/"))
            resp = session.post(url, headers=headers, data=body, timeout=timeout)

            if not resp.ok:
                # 部分学校会返回 HTML 登录页或错误页，这里截断打印便于排查
                preview = (resp.text or "")[:200]
                raise RuntimeError(f"HTTP {resp.status_code}，响应预览：{preview}")

            if not resp.content:
                raise RuntimeError("响应为空，可能未登录/被拦截/接口不可用")

            return resp.content

        except Exception as e:
            last_err = e

    raise RuntimeError(f"两条路径都失败，最后错误：{last_err}") from last_err

def fetch_input_times(
    base_url: str,
    xnm: str,
    xqm: str,
    cookie: str,
    timeout: int = 30,
) -> dict[str, str]:
    """
    获取成绩录入时间

    返回：
        {
            "教学班ID": "录入时间",
            ...
        }
    """

    session = requests.Session()

    url = urljoin(
        base_url.rstrip("/") + "/",
        "cjcx/cjcx_cxXsgrcj.html?doType=query"
    )

    headers = {
        "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8",
        "User-Agent": "Mozilla/5.0",
        "Cookie": cookie,
        "Referer": urljoin(base_url.rstrip("/") + "/", "cjcx/"),
    }

    data = {
        "xnm": xnm,
        "xqm": xqm,
        "queryModel.showCount": "1000",
        "queryModel.currentPage": "1",
        "queryModel.sortName": "",
        "queryModel.sortOrder": "asc",
    }

    resp = session.post(url, headers=headers, data=data, timeout=timeout)

    if not resp.ok:
        raise RuntimeError("获取成绩录入时间失败")

    js = resp.json()

    result = {}

    for item in js.get("items", []):
        jxb_id = item.get("jxb_id")
        tjsj = item.get("tjsj")

        if jxb_id and tjsj:
            result[jxb_id] = tjsj

    return result

# -----------------------------
# 3) 解析 Excel 二进制为 DataFrame / list[dict]
# -----------------------------
def parse_excel_bytes(content: bytes) -> pd.DataFrame:
    """
    将返回的 Excel 二进制解析成 DataFrame。
    先尝试 xlsx(openpyxl)，失败再尝试 xls(xlrd)。
    """
    bio = BytesIO(content)

    # 先按 xlsx
    try:
        df = pd.read_excel(bio, engine="openpyxl")
        return df
    except Exception:
        pass

    # 再按 xls
    bio.seek(0)
    df = pd.read_excel(bio, engine="xlrd")
    return df


def df_to_rows(df: pd.DataFrame) -> list[dict[str, Any]]:
    """
    DataFrame -> list[dict]，顺便做一点轻度清洗/容错。
    """
    df = df.copy()

    # 统一列名去空格（有些表头会有不可见空格）
    df.columns = [str(c).strip() for c in df.columns]

    # 仅保留我们关心的列（如果缺列就忽略）
    keep_cols = ["学年", "学期", "教学班ID", "学分", "课程名称", "成绩", "成绩分项"]
    cols = [c for c in keep_cols if c in df.columns]
    df = df[cols]

    # 把 NaN 转为 None
    df = df.where(pd.notnull(df), None)

    return df.to_dict(orient="records")


# -----------------------------
# 4) 合并同一门课 + 总评置顶
# -----------------------------
def merge_courses(
    rows: list[dict[str, Any]],
    input_times: dict[str, str] | None = None
) -> list[dict[str, Any]]:
    """
    把同一门课（同一教学班ID优先）合并为一条记录，并把“总评”放到成绩分项最前面。

    输出结构仍为 list[dict]，每条包含：
    - 学年/学期/教学班ID/学分/课程名称
    - 总评（单独字段）
    - 分项成绩（不含总评，保持原出现顺序）
    - 成绩分项（拼接字符串，且总评在最前）
    """
    groups: "OrderedDict[tuple, dict[str, Any]]" = OrderedDict()

    for r in rows:
        xnm = r.get("学年")
        xqm = r.get("学期")
        jxb_id = r.get("教学班ID")
        kcmc = r.get("课程名称")
        xf = r.get("学分")

        # 优先用教学班ID分组（最稳），否则退化到 (学年, 学期, 课程名称)
        key = (jxb_id,) if jxb_id else (xnm, xqm, kcmc)

        if key not in groups:
            groups[key] = {
                "学年": xnm,
                "学期": xqm,
                "教学班ID": jxb_id,
                "学分": xf,
                "课程名称": kcmc,
                "总评": None,
                "分项成绩": [],  # 仅放非“总评”
            }

        item_name = str(r.get("成绩分项") or "").strip()
        score = r.get("成绩")

        # 识别“总评”（可扩展：比如 "总评成绩"/"课程总评" 等）
        if item_name == "总评":
            groups[key]["总评"] = score
        else:
            groups[key]["分项成绩"].append({"名称": item_name, "成绩": score})

    merged: list[dict[str, Any]] = []
    for g in groups.values():
        parts: list[str] = []

        # 总评放最前
        if g["总评"] is not None:
            parts.append(f"总评:{g['总评']}")

        # 其余分项按原出现顺序
        for p in g["分项成绩"]:
            name = (p.get("名称") or "").strip() or "未知分项"
            parts.append(f"\n{name}:{p.get('成绩')}")
        parts.append("\n")

        jxb_id = g["教学班ID"]

        merged.append({
            "学年": g["学年"],
            "学期": g["学期"],
            "教学班ID": jxb_id,
            "学分": g["学分"],
            "课程名称": g["课程名称"],
            "总评": g["总评"],
            "成绩录入时间": input_times.get(jxb_id) if input_times else None,
            "分项成绩": g["分项成绩"],
            "成绩分项": "；".join(parts),
        })

    return merged


# -----------------------------
# 5) 打印输出
# -----------------------------
def print_merged(merged: list[dict[str, Any]]) -> None:
    """
    按课程打印合并后的结果
    """
    for c in merged:
        print(f"{c.get('课程名称')}:{c.get('总评')}")

    print()
    for c in merged:
        print(f"{c.get('课程名称')} {c.get("成绩分项")}学分:{c.get('学分')} 录入:{c.get('成绩录入时间')}\n")

# -----------------------------
# 入口：抓取 -> 解析 -> 合并 -> 打印
# -----------------------------
xqm_dict={
    "1":"3",
    "2":"12",
    "3":"16",
}

def main():
    base_url = os.environ.get("BASE_URL")
    sid = os.environ.get("SID")  # 学号
    password = os.environ.get("PASSWORD")  # 密码
    # xnm/xqm 必须与页面下拉框 value 一致
    xnm = os.environ.get("XNM")
    xqm = os.environ.get("XQM")
    if xqm in xqm_dict:
        xqm = xqm_dict[xqm]
    cookie_dict = login_and_get_cookie(
        base_url,
        sid,
        password,
    )
    cookie = "JSESSIONID="+cookie_dict.get("JSESSIONID")
    # print(cookie)
    # print(type(cookie))
    raw = fetch_export_bytes(base_url, xnm, xqm, cookie=cookie)
    input_times = fetch_input_times(base_url, xnm, xqm, cookie=cookie)
    df = parse_excel_bytes(raw)
    rows = df_to_rows(df)

    merged = merge_courses(rows, input_times)
    print(f"成绩数量：{len(merged)}\n")

    print_merged(merged)


if __name__ == "__main__":
    main()