# -*- coding: utf-8 -*-
"""
港股数据助手 - 腾讯云 SCF 云函数

功能: 把绿鞋数据(TradeGoMart)和港股行情(腾讯)转发给前端页面,
      解决浏览器跨域和接口不稳定问题。

部署方式: 在腾讯云 SCF 控制台创建"Web 函数",运行时选 Python 3.9+,
          直接把本文件作为入口代码上传,无需安装任何依赖。

本函数零第三方依赖(只用标准库),上传即可运行。
"""

import json
import urllib.request
import urllib.parse

# ============================
# 上游接口配置
# ============================
GREEN_API = "https://cloudapi.livereport8.com/livereport/GreenShoeTrace"
GREEN_HEADERS = {
    "Content-Type": "application/json",
    "User-Agent": "Mozilla/5.0",
    "lang": "chs",
    "Origin": "https://livereport.tradegomart.com",
    "Referer": "https://livereport.tradegomart.com/",
}
QUOTE_URL = "https://qt.gtimg.cn/q={codes}"
QUOTE_HEADERS = {
    "User-Agent": "Mozilla/5.0",
    "Referer": "https://gu.qq.com/",
}
TIMEOUT = 25

CORS_HEADERS = {
    "Content-Type": "application/json; charset=utf-8",
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type, lang",
    "Access-Control-Max-Age": "86400",
}


def make_response(body, status=200):
    """构造 SCF Web 函数 / API 网关兼容的响应"""
    if isinstance(body, (dict, list)):
        body = json.dumps(body, ensure_ascii=False)
    return {
        "isBase64Encoded": False,
        "statusCode": status,
        "headers": CORS_HEADERS,
        "body": body,
    }


def http_request(url, method="GET", payload=None, headers=None, timeout=TIMEOUT):
    """极简 HTTP 请求,只用标准库"""
    data = None
    if payload is not None:
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(url, data=data, headers=headers or {}, method=method)
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return resp.status, resp.read()


def proxy_greenshoe(name, payload):
    """转发 TradeGoMart 绿鞋接口"""
    url = f"{GREEN_API}/{name}"
    try:
        status, raw = http_request(url, method="POST", payload=payload or {},
                                   headers=GREEN_HEADERS)
        try:
            data = json.loads(raw.decode("utf-8"))
        except Exception:
            data = {"result": -1, "msg": "上游返回非JSON", "raw": raw.decode("utf-8", "ignore")[:500]}
        return make_response(data, status)
    except Exception as e:
        return make_response({"result": -1, "msg": f"上游请求失败: {e}"}, 502)


def proxy_quote(codes):
    """转发腾讯行情接口(返回 GBK 编码文本)"""
    url = QUOTE_URL.format(codes=urllib.parse.quote(codes))
    try:
        status, raw = http_request(url, headers=QUOTE_HEADERS)
        text = raw.decode("gbk", errors="replace")
        return make_response({"result": 1, "data": text}, status)
    except Exception as e:
        return make_response({"result": -1, "msg": f"行情请求失败: {e}"}, 502)


def main_handler(event, context):
    # 兼容 Web 函数与 API 网关两种事件格式
    if isinstance(event, dict):
        path = event.get("path", "") or ""
        method = event.get("httpMethod", "GET") or "GET"
        headers = event.get("headers", {}) or {}
        query = event.get("queryString", {}) or event.get("queryStringParameters", {}) or {}
        raw_body = event.get("body", "") or ""
        if event.get("isBase64Encoded"):
            import base64
            raw_body = base64.b64decode(raw_body).decode("utf-8", "ignore")
    else:
        path = ""
        method = "GET"
        headers = {}
        query = {}
        raw_body = ""

    # 预检请求直接放行
    if method == "OPTIONS":
        return make_response("")

    try:
        payload = json.loads(raw_body) if raw_body else {}
    except Exception:
        payload = {}

    path = path.strip()

    # 绿鞋股票名单
    if path in ("/api/greenshoe/list", "/api/greenshoe", ""):
        return proxy_greenshoe("greenShoeStockList", payload)

    # 绿鞋单股数据,例如 /api/greenshoe/06915
    if path.startswith("/api/greenshoe/"):
        name = path[len("/api/greenshoe/"):].strip("/")
        if not name:
            return make_response({"result": -1, "msg": "缺少股票名称"}, 400)
        return proxy_greenshoe(name, payload)

    # 行情,例如 /api/quote?codes=hk00700,hk09988
    if path == "/api/quote":
        codes = query.get("codes", "") or ""
        if not codes:
            return make_response({"result": -1, "msg": "缺少 codes 参数"}, 400)
        return proxy_quote(codes)

    # 健康检查
    if path == "/api/health":
        return make_response({"result": 1, "msg": "ok"})

    return make_response({"result": -1, "msg": f"未知路径: {path}"}, 404)
