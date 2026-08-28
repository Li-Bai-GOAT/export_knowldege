#!/usr/bin/env python3
# /// script
# requires-python = ">=3.10"
# dependencies = [
#     "playwright>=1.40",
#     "requests>=2.31",
#     "openpyxl>=3.1",
# ]
# ///
"""登录知识库管理系统，获取文件进度 ID，并复用旧脚本导出知识条目。

当前脚本实现：
1. 打开登录页并登录；
2. 进入知识库管理列表页；
3. 搜索一个知识库名称并进入详情页；
4. 从详情 URL 获取 knowledgeId；
5. 点击详情页“编辑”只读取向量模型，不提交任何修改；
6. 仅当向量模型为 text-embedding-v3 时，调用接口获取 knowledgeProcProgressId；
7. 复用 export_knowledge.py 的接口逻辑，按医院/知识库/接口文件名导出 xlsx。

账号、密码和医院名称从同目录 config.json 读取；也支持命令行参数和环境变量覆盖。

示例（Windows PowerShell）：
    python knowledge_query.py

首次运行前安装依赖：
    python -m pip install playwright
"""

from __future__ import annotations

import argparse
import getpass
import json
import os
import re
import sys
from pathlib import Path
from typing import Any, NoReturn
from urllib.parse import parse_qs, urlparse

import requests
from openpyxl import Workbook
from playwright.sync_api import Page, Request, Response, TimeoutError as PlaywrightTimeoutError, sync_playwright

from export_knowledge import build_rows, fetch_all

LOGIN_URL = "https://aicloud.eheren.com/heren/aimanagement/login"
KNOWLEDGE_LIST_URL = "https://aicloud.eheren.com/heren/aimanagement/knowledge/list"
ACCOUNT_PLACEHOLDER = "请输入账号"
PASSWORD_PLACEHOLDER = "请输入登录密码"
SEARCH_PLACEHOLDER = "输入知识库名称进行搜索"
LOGIN_BUTTON_NAME = "登 录"
DEFAULT_KNOWLEDGE_NAME = "医生信息知识库"
TARGET_VECTOR_MODEL = "text-embedding-v3"
VECTOR_MODEL_SELECTOR = ".hr-form-item__aiModelId input.hr-input__inner"
IMPORT_SUCCESS_STATUS = "导入成功"
PROGRESS_URL = "https://aicloud.eheren.com/ai-manager/knowledge/queryKnowledgeProcProgressList"
PROGRESS_PAGE_SIZE = 10
SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_CONFIG_PATH = SCRIPT_DIR / "config.json"
DEFAULT_OUTPUT_ROOT = SCRIPT_DIR / "output"
DEFAULT_KNOWLEDGE_NAMES = (
    "医生信息知识库",
    "科室信息知识库",
    "医院信息知识库",
    "医院地址知识库",
)


class KnowledgeNotFound(RuntimeError):
    """搜索结果中不存在指定知识库时使用的可恢复异常。"""


class VectorModelUnavailable(RuntimeError):
    """无法在只读编辑表单中确认向量模型时使用的异常。"""


def fail(message: str) -> NoReturn:
    print(f"错误: {message}", file=sys.stderr)
    raise SystemExit(1)


def load_config(config_path: Path) -> dict[str, Any]:
    if not config_path.exists():
        fail(f"未找到配置文件: {config_path}")
    try:
        with config_path.open("r", encoding="utf-8") as config_file:
            config = json.load(config_file)
    except json.JSONDecodeError as exc:
        fail(f"配置文件不是有效 JSON: {config_path} ({exc})")
    except OSError as exc:
        fail(f"无法读取配置文件: {config_path} ({exc})")
    if not isinstance(config, dict):
        fail("配置文件顶层必须是 JSON 对象")
    return config


def config_value(config: dict[str, Any], *keys: str, default: Any = None) -> Any:
    for key in keys:
        value = config.get(key)
        if value is not None and str(value).strip():
            return value
    return default


def get_credentials(args: argparse.Namespace, config: dict[str, Any]) -> tuple[str, str]:
    username = (
        args.username
        or config_value(config, "username", "userName")
        or os.getenv("HEREN_USERNAME")
        or input("登录账号: ").strip()
    )
    password = config_value(config, "password", "passWord") or os.getenv("HEREN_PASSWORD")
    if not password:
        password = getpass.getpass("登录密码: ")
    if not username:
        fail("账号不能为空")
    if not password:
        fail("密码不能为空")
    return username, password


def visible_body_text(page: Page, limit: int = 1200) -> str:
    return page.locator("body").inner_text(timeout=10_000)[:limit]


def capture_token_from_request(request: Request, token_holder: dict[str, str]) -> None:
    """捕获页面请求头中的 user-token，不输出、不落盘。"""
    try:
        headers = {key.lower(): value for key, value in request.headers.items()}
    except Exception:
        return
    token = headers.get("user-token")
    if token:
        token_holder["value"] = token


def capture_token_from_response(response: Response, token_holder: dict[str, str]) -> None:
    """兼容登录接口把 token 放在 JSON 响应体中的情况。"""
    if token_holder.get("value"):
        return
    if "/login" not in response.url.lower():
        return
    try:
        payload = response.json()
    except Exception:
        return

    def find_token(value: Any) -> str | None:
        if isinstance(value, dict):
            for key, item in value.items():
                if str(key).lower() in {"token", "usertoken", "user-token", "access_token", "accesstoken"}:
                    if isinstance(item, str) and item:
                        return item
                found = find_token(item)
                if found:
                    return found
        elif isinstance(value, list):
            for item in value:
                found = find_token(item)
                if found:
                    return found
        return None

    token = find_token(payload)
    if token:
        token_holder["value"] = token


def login(page: Page, username: str, password: str) -> None:
    page.goto(LOGIN_URL, wait_until="domcontentloaded", timeout=60_000)

    account = page.get_by_placeholder(ACCOUNT_PLACEHOLDER, exact=True)
    password_input = page.get_by_placeholder(PASSWORD_PLACEHOLDER, exact=True)
    try:
        account.wait_for(state="visible", timeout=20_000)
        password_input.wait_for(state="visible", timeout=20_000)
    except PlaywrightTimeoutError:
        fail(f"未找到登录表单，当前页面内容: {visible_body_text(page)}")

    account.fill(username)
    password_input.fill(password)

    captcha = page.get_by_placeholder("请输入验证码", exact=True)
    if captcha.count() > 0 and captcha.is_visible():
        fail("登录页出现验证码，请先手动完成验证码后再运行脚本")

    login_button = page.get_by_role("button", name=LOGIN_BUTTON_NAME, exact=True)
    if login_button.count() == 0:
        login_button = page.locator('button[type="submit"]')
    if login_button.count() == 0:
        fail("未找到登录按钮")

    login_button.first.click()
    try:
        page.wait_for_url(re.compile(r"/heren/aimanagement/(?:member/list|knowledge/list)"), timeout=60_000)
    except PlaywrightTimeoutError:
        fail(f"登录后未跳转到管理页面，当前地址: {page.url}，页面内容: {visible_body_text(page)}")


def search_knowledge(page: Page, knowledge_name: str) -> tuple[int, str]:
    page.goto(KNOWLEDGE_LIST_URL, wait_until="domcontentloaded", timeout=60_000)
    search_box = page.get_by_placeholder(SEARCH_PLACEHOLDER, exact=True)
    try:
        search_box.wait_for(state="visible", timeout=20_000)
    except PlaywrightTimeoutError:
        fail(f"未找到知识库搜索框，当前地址: {page.url}，页面内容: {visible_body_text(page)}")

    search_box.fill(knowledge_name)
    search_box.press("Enter")

    result = page.get_by_text(knowledge_name, exact=True)
    try:
        result.first.wait_for(state="visible", timeout=20_000)
    except PlaywrightTimeoutError:
        raise KnowledgeNotFound(f"搜索后未找到知识库: {knowledge_name}") from None
    matched_count = result.count()

    # 页面上的知识库卡片需要双击标题进入详情，普通单击不会跳转。
    title = page.locator(".agent-title").filter(has_text=knowledge_name)
    if title.count() == 0:
        raise KnowledgeNotFound(f"搜索结果存在，但未找到可双击的知识库标题: {knowledge_name}")

    detail_pattern = re.compile(r"/heren/aimanagement/knowledge/detail[?]knowledgeId=[^&#]+")
    title.first.dblclick()
    try:
        page.wait_for_url(detail_pattern, timeout=30_000)
    except PlaywrightTimeoutError:
        fail(f"双击知识库标题后未进入详情页，当前地址: {page.url}")

    knowledge_id = parse_qs(urlparse(page.url).query).get("knowledgeId", [None])[0]
    if not knowledge_id:
        fail(f"详情页 URL 中未找到 knowledgeId，当前地址: {page.url}")

    return matched_count, knowledge_id


def read_vector_model(page: Page) -> str:
    """打开编辑表单只读取向量模型；不点击确定、不提交任何修改。"""
    edit_button = page.get_by_text("编辑", exact=True)
    try:
        edit_button.first.wait_for(state="visible", timeout=20_000)
    except PlaywrightTimeoutError:
        raise VectorModelUnavailable(
            f"详情页未找到“编辑”按钮，当前地址: {page.url}"
        ) from None

    edit_button.first.click()
    model_input = page.locator(VECTOR_MODEL_SELECTOR).first
    try:
        model_input.wait_for(state="visible", timeout=20_000)
    except PlaywrightTimeoutError:
        raise VectorModelUnavailable(
            f"点击“编辑”后未找到向量模型字段，当前地址: {page.url}"
        ) from None

    vector_model = model_input.input_value().strip()
    if not vector_model:
        raise VectorModelUnavailable(
            f"向量模型字段为空，无法安全判断是否下载，当前地址: {page.url}"
        )
    return vector_model


PROGRESS_ID_KEYS = (
    "knowledgeProcProgressId",
    "knowledgeProcProgressID",
    "progressId",
    "progress_id",
)
FILE_ID_KEYS = ("fileId", "fileID", "file_id")
FILE_NAME_KEYS = ("fileName", "file_name", "originFileName", "originalFileName")


def first_value(record: dict[str, Any], keys: tuple[str, ...]) -> str | None:
    for key in keys:
        value = record.get(key)
        if value is not None and str(value).strip():
            return str(value).strip()
    return None


def find_progress_dicts(value: Any, found: list[dict[str, Any]]) -> None:
    if isinstance(value, dict):
        has_file_info = any(key in value for key in FILE_ID_KEYS + FILE_NAME_KEYS)
        has_progress_info = any(key in value for key in PROGRESS_ID_KEYS) or ("id" in value and has_file_info)
        if has_progress_info:
            found.append(value)
        for item in value.values():
            find_progress_dicts(item, found)
    elif isinstance(value, list):
        for item in value:
            find_progress_dicts(item, found)


def normalise_progress_records(payload: Any) -> list[dict[str, str | None]]:
    raw_records: list[dict[str, Any]] = []
    find_progress_dicts(payload, raw_records)
    records: list[dict[str, str | None]] = []
    seen: set[str] = set()
    for raw in raw_records:
        progress_id = first_value(raw, PROGRESS_ID_KEYS)
        if not progress_id and any(key in raw for key in FILE_ID_KEYS + FILE_NAME_KEYS):
            progress_id = first_value(raw, ("id",))
        if not progress_id or progress_id in seen:
            continue
        seen.add(progress_id)
        records.append(
            {
                "knowledgeProcProgressId": progress_id,
                "fileId": first_value(raw, FILE_ID_KEYS),
                "fileName": first_value(raw, FILE_NAME_KEYS),
                "processStatus": first_value(
                    raw,
                    (
                        "procStatusStr",
                        "processStatusStr",
                        "process_status_str",
                        "statusStr",
                        "status_str",
                        "processStatus",
                        "process_status",
                        "status",
                    ),
                ),
                "processStatusCode": first_value(
                    raw,
                    (
                        "procStatus",
                        "proc_status",
                        "processStatusCode",
                        "process_status_code",
                        "statusCode",
                        "status_code",
                    ),
                ),
            }
        )
    return records


def fetch_progress_records(knowledge_id: str, token: str) -> list[dict[str, str | None]]:
    headers = {"accept": "application/json", "user-token": token}
    records: list[dict[str, str | None]] = []
    seen: set[str] = set()

    for page_index in range(1, 101):
        response = requests.get(
            PROGRESS_URL,
            params={
                "knowledgeId": knowledge_id,
                "pageSize": PROGRESS_PAGE_SIZE,
                "pageIndex": page_index,
            },
            headers=headers,
            timeout=60,
        )
        if response.status_code == 403:
            fail("查询文件进度接口返回 403，当前登录 token 可能无效或已过期")
        response.raise_for_status()
        payload = response.json()
        page_records = normalise_progress_records(payload)
        for record in page_records:
            progress_id = record["knowledgeProcProgressId"]
            if progress_id and progress_id not in seen:
                seen.add(progress_id)
                records.append(record)

        if len(page_records) < PROGRESS_PAGE_SIZE:
            break

    return records


def safe_name(value: str, fallback: str) -> str:
    """生成 Windows 可用的目录名或文件名，避免接口返回非法字符。"""
    value = (value or fallback).strip()
    value = re.sub(r"[<>:\"/\\|?*\x00-\x1f]", "_", value)
    value = value.rstrip(" .")
    if not value:
        value = fallback
    if value.upper() in {"CON", "PRN", "AUX", "NUL"}:
        value = f"_{value}"
    return value


def get_output_path(
    output_root: Path,
    hospital_name: str,
    knowledge_name: str,
    file_name: str | None,
    progress_id: str,
    used_names: set[str],
) -> Path:
    """按 医院/知识库/接口文件名 生成唯一的 xlsx 输出路径。"""
    hospital_dir = safe_name(hospital_name, "未命名医院")
    knowledge_dir = safe_name(knowledge_name, "未命名知识库")
    source_name = (file_name or "").strip()
    source_name = re.split(r"[\\/]", source_name)[-1]
    source_stem = Path(source_name).stem if source_name else progress_id
    source_stem = safe_name(source_stem, progress_id)
    candidate_name = f"{source_stem}.xlsx"
    if candidate_name in used_names:
        candidate_name = f"{source_stem}_{safe_name(progress_id, progress_id)}.xlsx"
    used_names.add(candidate_name)
    return output_root / hospital_dir / knowledge_dir / candidate_name


def export_with_legacy_script(
    knowledge_id: str,
    progress_id: str,
    token: str,
    output_path: Path,
) -> tuple[str, int]:
    """复用旧脚本的 fetch_all/build_rows 逻辑，输出到指定的命名路径。"""
    items = fetch_all(knowledge_id, progress_id, token)
    if not items:
        return "", 0

    all_fields, rows = build_rows(items)
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "knowledge"
    worksheet.append(all_fields)
    for row in rows:
        worksheet.append([row.get(field, "") for field in all_fields])

    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)
    return str(output_path), len(rows)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="登录并搜索 Heren MindHub 知识库")
    parser.add_argument(
        "--config",
        default=str(DEFAULT_CONFIG_PATH),
        help=f"配置文件路径，默认: {DEFAULT_CONFIG_PATH}",
    )
    parser.add_argument("--username", help="登录账号；不传时读取 HEREN_USERNAME 或交互输入")
    parser.add_argument(
        "--knowledge-name",
        dest="knowledge_names",
        action="append",
        help="只处理指定知识库；可重复传入多次。不传时读取 config.json 中的 knowledgeNames",
    )
    parser.add_argument("--headless", action="store_true", help="无界面运行；默认打开可见的 Microsoft Edge")
    return parser.parse_args()


def get_knowledge_names(args: argparse.Namespace, config: dict[str, Any]) -> list[str]:
    configured = args.knowledge_names or config_value(config, "knowledgeNames", "knowledge_names")
    if configured is None:
        return list(DEFAULT_KNOWLEDGE_NAMES)
    if isinstance(configured, str):
        names = [configured.strip()]
    elif isinstance(configured, list):
        names = [str(item).strip() for item in configured if str(item).strip()]
    else:
        fail("knowledgeNames 必须是字符串或字符串数组")
    names = list(dict.fromkeys(name for name in names if name))
    if not names:
        fail("没有可处理的知识库名称")
    return names


def main() -> int:
    args = parse_args()
    config = load_config(Path(args.config).expanduser().resolve())
    username, password = get_credentials(args, config)
    knowledge_names = get_knowledge_names(args, config)
    hospital_name = str(config_value(config, "hospitalName", "hospital_name", default="demo"))
    output_root_value = config_value(config, "outputRoot", "output_root")
    output_root = Path(output_root_value).expanduser() if output_root_value else DEFAULT_OUTPUT_ROOT
    if not output_root.is_absolute():
        output_root = (SCRIPT_DIR / output_root).resolve()

    with sync_playwright() as playwright:
        try:
            browser = playwright.chromium.launch(channel="msedge", headless=args.headless)
        except Exception as exc:
            fail(
                "无法启动 Microsoft Edge。请确认已安装 Edge，并先执行: python -m pip install playwright"
                f"\n详细信息: {exc}"
            )

        page = browser.new_page(viewport={"width": 1440, "height": 900})
        token_holder: dict[str, str] = {}
        page.on("request", lambda request: capture_token_from_request(request, token_holder))
        page.on("response", lambda response: capture_token_from_response(response, token_holder))
        try:
            login(page, username, password)
            total_exported = 0
            summary: list[tuple[str, str, int, int, int]] = []

            for knowledge_name in knowledge_names:
                print(f"开始处理知识库: {knowledge_name}")
                try:
                    matched_count, knowledge_id = search_knowledge(page, knowledge_name)
                except KnowledgeNotFound as exc:
                    print(f"跳过知识库: {exc}")
                    summary.append((knowledge_name, "未找到", 0, 0, 0))
                    continue

                try:
                    vector_model = read_vector_model(page)
                except VectorModelUnavailable as exc:
                    fail(f"无法确认知识库“{knowledge_name}”的向量模型: {exc}")

                print(f"向量模型: {vector_model}")
                if vector_model.casefold() != TARGET_VECTOR_MODEL.casefold():
                    print(
                        f"跳过知识库: 向量模型为 {vector_model}，"
                        f"仅 {TARGET_VECTOR_MODEL} 允许下载"
                    )
                    summary.append((knowledge_name, "向量模型不匹配，已跳过", 0, 0, 0))
                    continue
                print(f"向量模型符合条件，继续下载: {TARGET_VECTOR_MODEL}")

                token = token_holder.get("value")
                if not token:
                    fail("登录成功但未捕获到页面请求中的 user-token，无法调用文件进度接口")

                try:
                    progress_records = fetch_progress_records(knowledge_id, token)
                except requests.RequestException as exc:
                    print(f"跳过知识库: 获取文件进度失败 ({exc})")
                    summary.append((knowledge_name, "获取文件进度失败", 0, 0, 0))
                    continue

                successful_records: list[dict[str, str | None]] = []
                for record in progress_records:
                    process_status = (record.get("processStatus") or "").strip()
                    if process_status == IMPORT_SUCCESS_STATUS:
                        successful_records.append(record)
                    else:
                        print(
                            f"跳过未导入成功文件: "
                            f"progressId={record.get('knowledgeProcProgressId') or '-'} "
                            f"| fileName={record.get('fileName') or '-'} "
                            f"| 状态={process_status or '未知'}"
                        )

                print("进入详情页并获取文件进度 ID 成功")
                print(f"详情页: {page.url}")
                print(f"知识库: {knowledge_name}")
                print(f"匹配到的同名文本数量: {matched_count}")
                print(f"knowledgeId: {knowledge_id}")
                print(f"文件进度记录数: {len(progress_records)}")
                print(f"导入成功记录数: {len(successful_records)}")

                exported_for_knowledge = 0
                used_names: set[str] = set()
                for record in successful_records:
                    progress_id = record["knowledgeProcProgressId"]
                    if not progress_id:
                        continue
                    target_path = get_output_path(
                        output_root,
                        hospital_name,
                        knowledge_name,
                        record.get("fileName"),
                        progress_id,
                        used_names,
                    )
                    try:
                        output_path, row_count = export_with_legacy_script(
                            knowledge_id,
                            progress_id,
                            token,
                            target_path,
                        )
                    except requests.RequestException as exc:
                        print(f"导出失败: progressId={progress_id} ({exc})")
                        continue
                    if output_path:
                        exported_for_knowledge += 1
                        total_exported += 1
                        print(
                            f"导出完成: {output_path}"
                            f" | fileName={record.get('fileName') or '-'}"
                            f" | fileId={record.get('fileId') or '-'}"
                            f" | 数据条数={row_count}"
                        )
                    else:
                        print(f"跳过空文件: progressId={progress_id}")

                summary.append(
                    (
                        knowledge_name,
                        "完成",
                        len(progress_records),
                        len(successful_records),
                        exported_for_knowledge,
                    )
                )

            print("\n本次运行汇总:")
            for knowledge_name, status, progress_count, successful_count, exported_count in summary:
                print(
                    f"- {knowledge_name}: {status}；"
                    f"文件记录 {progress_count} 条，导入成功 {successful_count} 条，"
                    f"成功导出 {exported_count} 个文件"
                )
            if total_exported == 0:
                fail("四个知识库均没有导出任何数据")
            return 0
        finally:
            browser.close()


if __name__ == "__main__":
    raise SystemExit(main())
