#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""edsdk 长超时调用：doc_to_image 需要上传大文件并等待转换。"""
import json
import sys
import urllib.request
import urllib.error

TOOL = sys.argv[1]
ARGS = json.loads(sys.argv[2])
TIMEOUT = int(sys.argv[3]) if len(sys.argv) > 3 else 600

OPENER = urllib.request.build_opener(urllib.request.ProxyHandler({}))
ENDPOINT = None
for port in range(39099, 39109):
    ep = f"http://127.0.0.1:{port}/mcp"
    try:
        data = json.dumps({"jsonrpc": "2.0", "id": 1, "method": "tools/list", "params": {}}).encode()
        req = urllib.request.Request(ep, data=data, headers={"Content-Type": "application/json"}, method="POST")
        with OPENER.open(req, timeout=2) as resp:
            obj = json.loads(resp.read().decode())
            if isinstance(obj.get("result"), dict) and isinstance(obj["result"].get("tools"), list):
                ENDPOINT = ep
                break
    except Exception:
        continue
if not ENDPOINT:
    print("error: editor_sdk not found")
    sys.exit(1)

data = json.dumps({"jsonrpc": "2.0", "id": 1, "method": "tools/call",
                   "params": {"name": TOOL, "arguments": ARGS}}).encode()
req = urllib.request.Request(ENDPOINT, data=data,
                             headers={"Content-Type": "application/json", "Accept": "application/json"},
                             method="POST")
try:
    with OPENER.open(req, timeout=TIMEOUT) as resp:
        obj = json.loads(resp.read().decode())
except Exception as e:
    print(f"error: {e}")
    sys.exit(1)
if obj.get("error"):
    print(f"rpc error: {obj['error']}")
    sys.exit(1)
result = obj.get("result") or {}
texts = [c.get("text", "") for c in (result.get("content") or []) if isinstance(c, dict) and c.get("type") == "text"]
print("\n".join(texts) if texts else json.dumps(result, ensure_ascii=False, indent=2))
