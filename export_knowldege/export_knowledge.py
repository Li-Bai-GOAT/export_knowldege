#!/usr/bin/env -S uv run --script
# /// script
# requires-python = ">=3.10"
# dependencies = [
#     "requests",
#     "openpyxl",
# ]
# ///
"""导出知识库文档列表到 xlsx。

用法:
    uv run --script export_knowledge.py --knowledgeId 1896734298707464195 --knowledgeProcProgressId 4995173832232800257 --token 你的token

接口需要 token 鉴权（请求头 user-token）。token 从浏览器登录后该接口的请求头复制。
"""
import argparse
import os
import re
import sys

import requests
from openpyxl import Workbook

API_URL = "https://aicloud.eheren.com/ai-manager/knowledge/queryKnowOperateDocumentList"
LIMIT = 20
META_EXCLUDE = {"enable", "progress_id", "sort_num"}
OUT_DIR = r"C:\Users\ASUS\Desktop\export_knowldege\export_knowldege\output"


def parse_document_content(content: str | dict) -> dict:
    """把 --xxx--\n值 ... 解析成 {xxx: 值}。

    字段值为该 --xxx-- 标记之后、到下一个 --xxx-- 标记之前的内容。
    """
    if not content:
        return {}
    if isinstance(content, dict):
        return content
    if not isinstance(content, str):
        content = str(content)
    matches = list(re.finditer(r"--(.+?)--", content))
    if not matches:
        return {}
    result = {}
    for i, m in enumerate(matches):
        name = m.group(1).strip()
        start = m.end()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(content)
        result[name] = content[start:end].strip()
    return result


def fetch_all(knowledge_id: str, progress_id: str, token: str | None) -> list:
    all_items = []
    next_id = None
    page = 0
    while True:
        params = {
            "knowledgeId": knowledge_id,
            "knowledgeProcProgressId": progress_id,
            "limit": LIMIT,
        }
        if next_id:
            params["nextId"] = next_id
        headers = {"accept": "application/json"}
        if token:
            headers["user-token"] = token
        resp = requests.get(API_URL, params=params, headers=headers, timeout=60)
        if resp.status_code == 403:
            body = resp.text[:500]
            print(f"鉴权失败 (403): {body}", file=sys.stderr)
            print("可能原因: token 已过期，请到浏览器重新登录并复制最新的 user-token 请求头", file=sys.stderr)
            sys.exit(2)
        resp.raise_for_status()
        data = resp.json()
        # 接口实际数据在 data 字段里
        data = data.get("data") if isinstance(data, dict) and "data" in data else data
        items = data.get("knowledgeItemVos") or []
        all_items.extend(items)
        page += 1
        print(f"第 {page} 页拉取 {len(items)} 条，累计 {len(all_items)} 条", file=sys.stderr)
        next_id = data.get("nextId")
        if not next_id:
            break
    return all_items


def build_rows(items: list) -> tuple[list[str], list[dict]]:
    meta_fields: list[str] = []
    doc_fields: list[str] = []
    rows: list[dict] = []
    for item in items:
        meta = item.get("metadata") or {}
        doc = parse_document_content(item.get("documentContent") or "")
        row: dict = {}
        for k, v in meta.items():
            if k in META_EXCLUDE:
                continue
            field = f"{k}:META"
            if field not in meta_fields:
                meta_fields.append(field)
            row[field] = v
        for k, v in doc.items():
            if k not in doc_fields:
                doc_fields.append(k)
            row[k] = v
        rows.append(row)
    return meta_fields + doc_fields, rows


def main() -> int:
    p = argparse.ArgumentParser(description="导出知识库文档列表到 xlsx")
    p.add_argument("--knowledgeId", required=True, help="知识库 ID")
    p.add_argument("--knowledgeProcProgressId", required=True, help="进度 ID，输出文件名也用它")
    p.add_argument("--token", default=None, help="登录 token，作为 user-token 请求头发送（必需，从浏览器请求头复制）")
    args = p.parse_args()

    if not args.token:
        print("错误: 需要 --token 进行鉴权", file=sys.stderr)
        return 2

    items = fetch_all(args.knowledgeId, args.knowledgeProcProgressId, args.token)
    if not items:
        print("未拉取到任何数据", file=sys.stderr)
        return 1

    all_fields, rows = build_rows(items)

    wb = Workbook()
    ws = wb.active
    ws.title = "knowledge"
    ws.append(all_fields)
    for row in rows:
        ws.append([row.get(f, "") for f in all_fields])

    os.makedirs(OUT_DIR, exist_ok=True)
    out_path = os.path.join(OUT_DIR, f"{args.knowledgeProcProgressId}.xlsx")
    wb.save(out_path)
    print(f"导出完成: {out_path}\n共 {len(rows)} 条数据，{len(all_fields)} 个字段")
    print(f"字段: {all_fields}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
