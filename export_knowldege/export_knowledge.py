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
import logging
import os
import re
from pathlib import Path

import requests
from openpyxl import Workbook

from runtime_logging import configure_logging

API_URL = "https://aicloud.eheren.com/ai-manager/knowledge/queryKnowOperateDocumentList"
LIMIT = 20
META_EXCLUDE = {"enable", "progress_id", "sort_num"}
OUT_DIR = r"C:\Users\ASUS\Desktop\export_knowldege\export_knowldege\output"
DEFAULT_LOG_DIR = Path(OUT_DIR) / "logs"
logger = logging.getLogger("export_knowledge")


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
    logger.info("开始读取知识条目: knowledgeId=%s, progressId=%s", knowledge_id, progress_id)
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
            logger.error("鉴权失败 (403): knowledgeId=%s, progressId=%s", knowledge_id, progress_id)
            logger.error("可能原因: token 已过期，请到浏览器重新登录并复制最新的 user-token 请求头")
            raise SystemExit(2)
        resp.raise_for_status()
        data = resp.json()
        # 接口实际数据在 data 字段里
        data = data.get("data") if isinstance(data, dict) and "data" in data else data
        items = data.get("knowledgeItemVos") or []
        all_items.extend(items)
        page += 1
        logger.info("第 %s 页拉取 %s 条，累计 %s 条", page, len(items), len(all_items))
        next_id = data.get("nextId")
        if not next_id:
            break
    logger.info("知识条目读取完成: knowledgeId=%s, progressId=%s, totalItems=%s", knowledge_id, progress_id, len(all_items))
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
    p.add_argument("--log-dir", default=str(DEFAULT_LOG_DIR), help=f"日志目录，默认: {DEFAULT_LOG_DIR}")
    args = p.parse_args()
    global logger
    logger, log_path = configure_logging(Path(args.log_dir).expanduser().resolve(), "export_knowledge")
    logger.info("开始手动导出: knowledgeId=%s, progressId=%s", args.knowledgeId, args.knowledgeProcProgressId)

    if not args.token:
        logger.error("错误: 需要 --token 进行鉴权")
        return 2

    try:
        items = fetch_all(args.knowledgeId, args.knowledgeProcProgressId, args.token)
        if not items:
            logger.warning("未拉取到任何数据")
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
        logger.info("导出完成: %s | 共 %s 条数据，%s 个字段", out_path, len(rows), len(all_fields))
        logger.info("字段: %s", all_fields)
        logger.info("本次运行结束，日志文件: %s", log_path)
        return 0
    except Exception:
        logger.exception("手动导出发生未预期异常")
        raise


if __name__ == "__main__":
    raise SystemExit(main())
