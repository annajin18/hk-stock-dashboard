# -*- coding: utf-8 -*-
"""
港股数据助手 - 腾讯云 SCF Web 函数

功能: 把绿鞋数据(TradeGoMart)和港股行情(腾讯)转发给前端页面,
      解决浏览器跨域和接口不稳定问题。

部署方式: 在腾讯云 SCF 控制台创建"Web 函数",运行时选 Python 3.9+,
          上传本目录打包的 hk-data-proxy.zip(包含 index.py 和 scf_bootstrap)。
          本函数零第三方依赖(只用标准库)。
"""

import json
import os
import time
import urllib.request
import urllib.parse
from http.server import BaseHTTPRequestHandler, HTTPServer

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

# 页面托管：把 COS 上的静态页面通过云函数输出，解决 COS 默认域名强制下载问题
PAGE_BASE = "https://hk-dashboard-1462592466.cos.ap-guangzhou.myqcloud.com"
PAGE_FILES = {
    "/": "index.html",
    "/index.html": "index.html",
    "/hk_greenshoe.html": "hk_greenshoe.html",
    "/greenshoe_detail.html": "greenshoe_detail.html",
    "/hk_unlock_overview.html": "hk_unlock_overview.html",
    "/hk_ggt_ru_tong.html": "hk_ggt_ru_tong.html",
    "/hk_ipo_reminder.html": "hk_ipo_reminder.html",
    "/crs_calculator.html": "crs_calculator.html",
}
PAGE_CACHE = {}
PAGE_CACHE_TTL = 60  # 秒
PAGE_HTML_HEADERS = {
    "Content-Type": "text/html; charset=utf-8",
    "Access-Control-Allow-Origin": "*",
    "Cache-Control": "public, max-age=60",
}

CORS_HEADERS = {
    "Content-Type": "application/json; charset=utf-8",
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type, lang",
    "Access-Control-Max-Age": "86400",
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


def make_response(body, status=200):
    """统一响应结构"""
    if isinstance(body, (dict, list)):
        body = json.dumps(body, ensure_ascii=False)
    return status, body


def fetch_page_file(fname):
    """从 COS 拉取页面文件,返回文本;失败返回 None"""
    url = PAGE_BASE + "/" + fname
    try:
        status, raw = http_request(url, headers={"User-Agent": "Mozilla/5.0"})
        if status == 200:
            return raw.decode("utf-8", "replace")
    except Exception:
        pass
    return None


def proxy_page(path):
    """把数据库的页面转发给浏览器,使用正常 text/html 返回"""
    fname = PAGE_FILES.get(path)
    if not fname:
        return None
    now = time.time()
    cached = PAGE_CACHE.get(path)
    if cached and now - cached[0] < PAGE_CACHE_TTL:
        return 200, cached[1], PAGE_HTML_HEADERS
    body = fetch_page_file(fname)
    if body is None:
        return None
    PAGE_CACHE[path] = (now, body)
    return 200, body, PAGE_HTML_HEADERS


def proxy_vendor(path):
    """把 /vendor/ 下的 JS/CSS 资源从 COS 转发给前端,不受下载头影响"""
    fname = path[len("/vendor/"):]
    if not fname or ".." in fname:
        return None
    url = PAGE_BASE + "/vendor/" + fname
    try:
        status, raw = http_request(url, headers={"User-Agent": "Mozilla/5.0"})
        if status == 200:
            ctype = "application/javascript; charset=utf-8" if fname.endswith(".js") else "text/plain; charset=utf-8"
            headers = {
                "Content-Type": ctype,
                "Access-Control-Allow-Origin": "*",
                "Cache-Control": "public, max-age=300",
            }
            return 200, raw.decode("utf-8", "replace"), headers
    except Exception:
        pass
    return None


def route(method, path, query, raw_body):
    """路由分发,返回 (status, body, headers)"""
    if method == "OPTIONS":
        return 200, "", CORS_HEADERS

    # 页面代理:访问网址直接返回 HTML 页面
    if method == "GET":
        page = proxy_page(path)
        if page:
            return page
        if path.startswith("/vendor/"):
            vendor = proxy_vendor(path)
            if vendor:
                return vendor

    try:
        payload = json.loads(raw_body) if raw_body else {}
    except Exception:
        payload = {}

    path = path.strip()

    # 绿鞋股票名单
    if path in ("/api/greenshoe/list", "/api/greenshoe"):
        status, body = proxy_greenshoe("greenShoeStockList", payload)
        return status, body, CORS_HEADERS

    # 绿鞋单股数据,例如 /api/greenshoe/06915
    if path.startswith("/api/greenshoe/"):
        name = path[len("/api/greenshoe/"):].strip("/")
        if not name:
            return 400, json.dumps({"result": -1, "msg": "缺少股票名称"}, ensure_ascii=False), CORS_HEADERS
        status, body = proxy_greenshoe(name, payload)
        return status, body, CORS_HEADERS

    # 行情,例如 /api/quote?codes=hk00700,hk09988
    if path == "/api/quote":
        codes = (query.get("codes") or [""])[0]
        if not codes:
            return 400, json.dumps({"result": -1, "msg": "缺少 codes 参数"}, ensure_ascii=False), CORS_HEADERS
        status, body = proxy_quote(codes)
        return status, body, CORS_HEADERS

    # 健康检查
    if path == "/api/health":
        return 200, json.dumps({"result": 1, "msg": "ok"}, ensure_ascii=False), CORS_HEADERS

    return 404, json.dumps({"result": -1, "msg": f"未知路径: {path}"}, ensure_ascii=False), CORS_HEADERS


class Handler(BaseHTTPRequestHandler):
    """Web 函数 HTTP 服务处理器"""

    def _handle(self):
        try:
            path = self.path.split("?", 1)[0]
            query = {}
            if "?" in self.path:
                parsed = urllib.parse.parse_qs(self.path.split("?", 1)[1])
                query = {k: v for k, v in parsed.items()}
            length = int(self.headers.get("Content-Length") or 0)
            raw_body = self.rfile.read(length).decode("utf-8", "ignore") if length else ""
            status, body, headers = route(self.command, path, query, raw_body)
            self.send_response(status)
            for k, v in headers.items():
                self.send_header(k, v)
            self.end_headers()
            self.wfile.write(body.encode("utf-8"))
        except Exception as e:
            try:
                self.send_response(500)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({"result": -1, "msg": f"内部错误: {e}"},
                                            ensure_ascii=False).encode("utf-8"))
            except Exception:
                pass

    def do_GET(self):
        self._handle()

    def do_POST(self):
        self._handle()

    def do_OPTIONS(self):
        self.send_response(200)
        for k, v in CORS_HEADERS.items():
            self.send_header(k, v)
        self.end_headers()

    def log_message(self, fmt, *args):
        pass


def main_handler(event, context):
    """兼容事件函数调用方式(API 网关触发),便于本地测试"""
    if isinstance(event, dict):
        path = event.get("path", "") or ""
        method = event.get("httpMethod", "GET") or "GET"
        query = event.get("queryString", {}) or event.get("queryStringParameters", {}) or {}
        raw_body = event.get("body", "") or ""
        if event.get("isBase64Encoded"):
            import base64
            raw_body = base64.b64decode(raw_body).decode("utf-8", "ignore")
    else:
        path, method, query, raw_body = "", "GET", {}, ""
    status, body, _ = route(method, path, query, raw_body)
    return {
        "isBase64Encoded": False,
        "statusCode": status,
        "headers": CORS_HEADERS,
        "body": body,
    }


if __name__ == "__main__":
    port = int(os.environ.get("SCF_RUNTIME_PORT", "9000"))
    server = HTTPServer(("0.0.0.0", port), Handler)
    print(f"SCF Web function listening on {port}")
    server.serve_forever()
