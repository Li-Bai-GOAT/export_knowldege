#!/usr/bin/env python3
# /// script
# requires-python = ">=3.10"
# dependencies = [
#     "playwright>=1.40",
#     "requests>=2.31",
#     "openpyxl>=3.1",
# ]
# ///
"""登录知识库管理系统，导出源知识库并可选重建、上传 `_bge` 文件。

当前脚本实现：
1. 打开登录页并登录；
2. 进入知识库管理列表页；
3. 搜索一个知识库名称并进入详情页；
4. 从详情 URL 获取 knowledgeId；
5. 点击详情页“编辑”只读记录全部配置字段，不提交任何修改；
6. 调用接口记录知识库文件列表及文件状态；
7. 仅当向量模型为 text-embedding-v3 时，调用接口获取 knowledgeProcProgressId；
8. 复用 export_knowledge.py 的接口逻辑，按医院/知识库/接口文件名导出 xlsx；
9. 可选通过右上角 UI 创建 `_bge` 知识库，并逐文件上传导出的 Excel；
10. 可选持续轮询 `_bge` 文件导入状态并保存结果。

账号、密码和医院名称从同目录 config.json 读取；也支持命令行参数和环境变量覆盖。

示例（Windows PowerShell）：
    python knowledge_query.py

首次运行前安装依赖：
    python -m pip install playwright
"""

from __future__ import annotations

import argparse
from datetime import datetime
import getpass
import json
import logging
import os
import re
import sys
import time
from pathlib import Path
from typing import Any, NoReturn
from urllib.parse import parse_qs, urlparse

import requests
from openpyxl import Workbook
from playwright.sync_api import Page, Request, Response, TimeoutError as PlaywrightTimeoutError, sync_playwright

from export_knowledge import build_rows, fetch_all
from runtime_logging import configure_logging

LOGIN_URL = "https://aicloud.eheren.com/heren/aimanagement/login"
KNOWLEDGE_LIST_URL = "https://aicloud.eheren.com/heren/aimanagement/knowledge/list"
ACCOUNT_PLACEHOLDER = "请输入账号"
PASSWORD_PLACEHOLDER = "请输入登录密码"
SEARCH_PLACEHOLDER = "输入知识库名称进行搜索"
LOGIN_BUTTON_NAME = "登 录"
DEFAULT_KNOWLEDGE_NAME = "医生信息知识库"
TARGET_VECTOR_MODEL = "text-embedding-v3"
BGE_VECTOR_MODEL = "bge-m3"
VECTOR_MODEL_SELECTOR = ".hr-form-item__aiModelId input.hr-input__inner"
IMPORT_SUCCESS_STATUS = "导入成功"
PROGRESS_URL = "https://aicloud.eheren.com/ai-manager/knowledge/queryKnowledgeProcProgressList"
DETAIL_URL = "https://aicloud.eheren.com/ai-manager/knowledge/getKnowledgeDetail"
DETAIL_PAGE_BASE_URL = "https://aicloud.eheren.com/heren/aimanagement/knowledge/detail"
UPLOAD_URL = "https://aicloud.eheren.com/ai-manager/knowledge/uploadKnowledgeFile"
PROC_STATUS_URL = "https://aicloud.eheren.com/ai-manager/commonData/queryKnowledgeProcStatus"
PROGRESS_PAGE_SIZE = 10
UPLOAD_POLL_TIMEOUT_SECONDS = 5 * 60
DROPDOWN_ACTION_TIMEOUT_MS = 2_500
IMPORT_SUCCESS_STATUS_CODES = {"2"}
SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_CONFIG_PATH = SCRIPT_DIR / "config.json"
DEFAULT_OUTPUT_ROOT = SCRIPT_DIR / "output"
DEFAULT_LOG_ROOT = SCRIPT_DIR / "logs"
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


class KnowledgeConfigurationUnavailable(RuntimeError):
    """源知识库编辑页字段不完整或无法安全复用时使用的异常。"""


class KnowledgeCreationUnavailable(RuntimeError):
    """创建或创建后复核失败时使用的可恢复异常。"""


class KnowledgeUploadUnavailable(RuntimeError):
    """目标知识库不存在、无上传权限或上传流程无法安全继续时使用的异常。"""


logger = logging.getLogger("knowledge_query")


def fail(message: str) -> NoReturn:
    logger.error("错误: %s", message)
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
    if token and not token_holder.get("value"):
        token_holder["value"] = token
        logger.info("已从页面请求捕获登录态 user-token（值已隐藏）")


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
    if token and not token_holder.get("value"):
        token_holder["value"] = token
        logger.info("已从登录响应捕获登录态 token（值已隐藏）")


def login(page: Page, username: str, password: str) -> None:
    logger.info("开始登录知识库管理系统: url=%s", LOGIN_URL)
    page.goto(LOGIN_URL, wait_until="domcontentloaded", timeout=60_000)
    logger.info("登录页面已打开")

    account = page.get_by_placeholder(ACCOUNT_PLACEHOLDER, exact=True)
    password_input = page.get_by_placeholder(PASSWORD_PLACEHOLDER, exact=True)
    try:
        account.wait_for(state="visible", timeout=20_000)
        password_input.wait_for(state="visible", timeout=20_000)
    except PlaywrightTimeoutError:
        fail(f"未找到登录表单，当前页面内容: {visible_body_text(page)}")

    account.fill(username)
    password_input.fill(password)
    logger.info("登录账号和密码已填入，准备提交")

    captcha = page.get_by_placeholder("请输入验证码", exact=True)
    if captcha.count() > 0 and captcha.is_visible():
        fail("登录页出现验证码，请先手动完成验证码后再运行脚本")

    login_button = page.get_by_role("button", name=LOGIN_BUTTON_NAME, exact=True)
    if login_button.count() == 0:
        login_button = page.locator('button[type="submit"]')
    if login_button.count() == 0:
        fail("未找到登录按钮")

    login_button.first.click()
    logger.info("登录表单已提交，等待管理页面")
    try:
        page.wait_for_url(re.compile(r"/heren/aimanagement/(?:member/list|knowledge/list)"), timeout=60_000)
    except PlaywrightTimeoutError:
        fail(f"登录后未跳转到管理页面，当前地址: {page.url}，页面内容: {visible_body_text(page)}")
    logger.info("登录成功，当前页面: %s", page.url)


def search_knowledge(page: Page, knowledge_name: str, result_index: int = 0) -> tuple[int, str]:
    logger.info("开始搜索知识库: name=%s, resultIndex=%s", knowledge_name, result_index)
    # 搜索列表和卡片均为异步渲染。偶发情况下页面已经打开，但第一次
    # 搜索没有等到接口返回，直接判断会误报“知识库不存在”。最多重新加载
    # 列表并重试一次，仍失败才视为未找到。
    last_error: str | None = None
    for attempt in range(2):
        logger.info("打开知识库列表页并执行搜索: name=%s, attempt=%s", knowledge_name, attempt + 1)
        page.goto(KNOWLEDGE_LIST_URL, wait_until="domcontentloaded", timeout=60_000)
        search_box = page.get_by_placeholder(SEARCH_PLACEHOLDER, exact=True)
        try:
            search_box.wait_for(state="visible", timeout=20_000)
            search_box.fill(knowledge_name)
            search_box.press("Enter")
            logger.info("知识库搜索条件已提交: name=%s, attempt=%s", knowledge_name, attempt + 1)

            result = page.get_by_text(knowledge_name, exact=True)
            result.first.wait_for(state="visible", timeout=20_000)
            break
        except PlaywrightTimeoutError:
            last_error = f"第 {attempt + 1} 次搜索未出现结果"
            if attempt == 0:
                logger.warning(
                    "搜索知识库结果未及时出现，将重新加载列表重试一次: %s",
                    knowledge_name,
                )
                continue
            raise KnowledgeNotFound(
                f"搜索后未找到知识库: {knowledge_name}（{last_error}）"
            ) from None
    else:
        raise KnowledgeNotFound(f"搜索后未找到知识库: {knowledge_name}") from None

    # 页面上的知识库卡片需要双击标题进入详情，普通单击不会跳转。
    title = page.locator(".agent-title").filter(has_text=knowledge_name)
    if title.count() == 0:
        logger.warning("搜索结果中未找到可双击的知识库标题: %s", knowledge_name)
        raise KnowledgeNotFound(f"搜索结果存在，但未找到可双击的知识库标题: {knowledge_name}")
    matched_count = title.count()
    logger.info("知识库搜索成功: name=%s, matchedCount=%s", knowledge_name, matched_count)

    if result_index < 0 or result_index >= title.count():
        raise KnowledgeNotFound(
            f"知识库结果序号超出范围: {knowledge_name} (index={result_index}, count={title.count()})"
        )

    detail_pattern = re.compile(r"/heren/aimanagement/knowledge/detail[?]knowledgeId=[^&#]+")
    title.nth(result_index).dblclick()
    try:
        page.wait_for_url(detail_pattern, timeout=30_000)
    except PlaywrightTimeoutError:
        fail(f"双击知识库标题后未进入详情页，当前地址: {page.url}")

    knowledge_id = parse_qs(urlparse(page.url).query).get("knowledgeId", [None])[0]
    if not knowledge_id:
        fail(f"详情页 URL 中未找到 knowledgeId，当前地址: {page.url}")

    logger.info("已进入知识库详情页: name=%s, knowledgeId=%s", knowledge_name, knowledge_id)
    return matched_count, knowledge_id


def redact_sensitive(value: Any) -> Any:
    """对快照中的敏感字段做脱敏，避免把 token、密码或密钥落盘。"""
    sensitive_words = ("token", "password", "passwd", "secret", "apikey", "api_key", "accesskey")
    if isinstance(value, dict):
        result: dict[str, Any] = {}
        for key, item in value.items():
            key_text = str(key)
            normalized_key = key_text.casefold().replace("-", "").replace(" ", "")
            if any(word.replace("_", "") in normalized_key for word in sensitive_words):
                result[key_text] = "***REDACTED***"
            else:
                result[key_text] = redact_sensitive(item)
        return result
    if isinstance(value, list):
        return [redact_sensitive(item) for item in value]
    return value


def _close_open_form(page: Page) -> None:
    """关闭当前打开的编辑/创建表单，只允许点击取消或按 Escape。"""
    try:
        open_form = page.locator(
            ".hr-drawer.hr-drawer--open, .hr-dialog, .hr-modal, [role='dialog']"
        ).last
        cancel_button = open_form.get_by_role("button", name="取消", exact=True)
        if cancel_button.count() > 0 and cancel_button.last.is_visible():
            cancel_button.last.click()
            return
    except Exception:
        pass
    try:
        page.keyboard.press("Escape")
    except Exception:
        pass


def _edit_field_value(field: Any) -> dict[str, Any]:
    """从编辑页字段对象中取出可复用的值和显示文本。"""
    if not isinstance(field, dict):
        return {"value": None, "text": None}
    value = field.get("value")
    text = field.get("text")
    return {
        "value": str(value).strip() if value is not None and str(value).strip() else None,
        "text": str(text).strip() if text is not None and str(text).strip() else None,
    }


def _edit_field_text(field: Any) -> str:
    """优先使用控件显示文本，兼容自定义下拉框和普通 input。"""
    normalized = _edit_field_value(field)
    return normalized.get("text") or normalized.get("value") or ""


def normalize_edit_config(config: Any) -> dict[str, Any]:
    """把编辑页 DOM 快照规范化为创建新知识库所需的字段。"""
    if not isinstance(config, dict):
        raise KnowledgeConfigurationUnavailable("编辑页未返回可识别的配置对象")

    normalized = {
        "name": _edit_field_text(config.get("name")),
        "code": _edit_field_text(config.get("code")),
        "description": _edit_field_text(config.get("description")),
        "scope": _edit_field_value(config.get("scope")),
        "private_type": config.get("private_type"),
        "knowledge_molds": config.get("knowledge_molds") or [],
        "vector_model": _edit_field_text(config.get("vector_model")),
        "use_visual_model": config.get("use_visual_model"),
        "knowledge_class": _edit_field_value(config.get("knowledge_class")),
    }
    if not isinstance(normalized["private_type"], dict):
        normalized["private_type"] = None
    if not isinstance(normalized["knowledge_molds"], list):
        normalized["knowledge_molds"] = []
    normalized["knowledge_molds"] = [
        option for option in normalized["knowledge_molds"] if isinstance(option, dict)
    ]

    required = {
        "name": normalized["name"],
        "code": normalized["code"],
        "scope": _edit_field_text(normalized["scope"]),
        "vector_model": normalized["vector_model"],
        "knowledge_class": _edit_field_text(normalized["knowledge_class"]),
    }
    missing = [key for key, value in required.items() if not value]
    if not normalized["private_type"]:
        missing.append("private_type")
    if not normalized["knowledge_molds"]:
        missing.append("knowledge_molds")
    if not isinstance(normalized["use_visual_model"], bool):
        missing.append("use_visual_model")
    if missing:
        raise KnowledgeConfigurationUnavailable(
            "编辑页字段缺失或无法识别: " + ", ".join(dict.fromkeys(missing))
        )
    return normalized


def collect_edit_form(page: Page) -> tuple[str, list[dict[str, Any]], dict[str, Any]]:
    """打开编辑表单，采集原始和规范化字段后关闭；绝不点击确定。"""
    logger.info("开始只读采集知识库编辑配置: page=%s", page.url)
    edit_button = page.get_by_text("编辑", exact=True)
    try:
        edit_button.first.wait_for(state="visible", timeout=20_000)
    except PlaywrightTimeoutError:
        raise VectorModelUnavailable(
            f"详情页未找到“编辑”按钮，当前地址: {page.url}"
        ) from None

    edit_button.first.click()
    logger.info("已打开知识库编辑表单")
    dialog = page.locator(
        ".hr-drawer.hr-drawer--open, .hr-dialog, .hr-modal, [role='dialog']"
    ).last
    model_input = page.locator(VECTOR_MODEL_SELECTOR).last
    try:
        dialog.wait_for(state="visible", timeout=20_000)
        model_input.wait_for(state="visible", timeout=20_000)
    except PlaywrightTimeoutError:
        _close_open_form(page)
        raise VectorModelUnavailable(
            f"点击“编辑”后未找到向量模型字段，当前地址: {page.url}"
        ) from None

    try:
        vector_model = model_input.input_value().strip()
        if not vector_model:
            raise VectorModelUnavailable(
                f"向量模型字段为空，无法安全判断是否下载，当前地址: {page.url}"
            )
        form_items = dialog.locator(".hr-form__item, .hr-form-item")
        if form_items.count() == 0:
            form_items = page.locator(".hr-form__item, .hr-form-item")
        edit_form: list[dict[str, Any]] = form_items.evaluate_all("""(items) => items.map((item, index) => {
        const text = (node) => (node?.innerText || node?.textContent || '').trim();
        const attrs = (node) => {
            if (!node) return null;
            const result = {};
            for (const attr of node.attributes || []) result[attr.name] = attr.value;
            return result;
        };
        const controls = Array.from(item.querySelectorAll('input, textarea, select, button, [role="combobox"], [role="radio"], [role="checkbox"]'))
            .map((control) => ({
                tag: control.tagName.toLowerCase(),
                type: control.getAttribute('type'),
                value: control.value ?? control.getAttribute('value'),
                text: text(control),
                checked: Boolean(control.checked),
                disabled: Boolean(control.disabled),
                readOnly: Boolean(control.readOnly),
                placeholder: control.getAttribute('placeholder'),
                name: control.getAttribute('name'),
                id: control.getAttribute('id'),
                dataValue: control.getAttribute('data-value'),
                className: control.className?.toString() || '',
                attributes: attrs(control)
            }));
        const labelNode = item.querySelector('.hr-form-item__label, .hr-form__label, label');
        const keyNode = item.querySelector('[name], [id], [data-field], [data-key]');
        return {
            index,
            fieldKey: keyNode?.getAttribute('name') || keyNode?.getAttribute('id') || keyNode?.getAttribute('data-field') || keyNode?.getAttribute('data-key') || null,
            label: text(labelNode),
            text: text(item),
            attributes: attrs(item),
            controls
        };
    })""")

        normalized_dom = dialog.evaluate(r"""(root) => {
            const text = (node) => (node?.innerText || node?.textContent || '').trim();
            const clean = (value) => {
                const result = value == null ? '' : String(value).trim();
                return result || null;
            };
            const items = Array.from(root.querySelectorAll('.hr-form-item, .hr-form__item'));
            const labelText = (item) => text(item.querySelector(
                '.hr-form-item__label, .hr-form__label, label'
            ));
            const findItem = (selector, labels) => {
                const direct = selector ? root.querySelector(selector) : null;
                if (direct) return direct.closest('.hr-form-item, .hr-form__item') || direct;
                return items.find((item) => labels.some((label) =>
                    labelText(item).replace(/\s/g, '').includes(label)
                )) || null;
            };
            const readField = (selector, labels) => {
                const item = findItem(selector, labels);
                if (!item) return {value: null, text: null};
                const control = item.querySelector(
                    'input:not([type="radio"]):not([type="checkbox"]), textarea, select'
                );
                let value = control?.value ?? null;
                let selectedText = null;
                if (control?.tagName?.toLowerCase() === 'select') {
                    selectedText = control.selectedOptions?.[0]?.textContent?.trim() || null;
                }
                if (!selectedText) {
                    const selected = item.querySelector(
                        '[aria-selected="true"], .hr-select__selected, .hr-select__selection,' +
                        ' .hr-input__inner, .hr-select__selected-item'
                    );
                    selectedText = text(selected) || null;
                }
                // 有控件但值为空时保持为空，不能把“知识库说明”等字段标签
                // 当成字段内容保存。
                return {value: clean(value), text: clean(selectedText || value || (control ? null : text(item)))};
            };
            const readRadio = (selector, labels) => {
                const item = findItem(selector, labels);
                if (!item) return null;
                const checked = item.querySelector('input[type="radio"]:checked');
                if (!checked) return null;
                const label = checked.closest('label') || checked.parentElement;
                return {value: clean(checked.value), text: clean(text(label) || checked.value)};
            };
            const readCheckboxes = (selector, labels) => {
                const item = findItem(selector, labels);
                if (!item) return [];
                return Array.from(item.querySelectorAll('input[type="checkbox"]')).map((control) => {
                    const label = control.closest('label') || control.parentElement;
                    return {
                        value: clean(control.value),
                        text: clean(text(label) || control.value),
                        checked: Boolean(control.checked),
                        disabled: Boolean(control.disabled)
                    };
                });
            };
            const visualItem = findItem('.hr-form-item__useVLModel', ['启动视觉模型']);
            return {
                name: readField('.hr-form-item__knowledgeTypeName', ['知识库名称']),
                code: readField('.hr-form-item__knowledgeTypeCode', ['知识库编码']),
                description: readField('.hr-form-item__knowledgeTypeDesc', ['知识库说明']),
                scope: readField('.hr-form-item__knowledgeScap', ['使用范围']),
                private_type: readRadio(null, ['私有库类型']),
                knowledge_molds: readCheckboxes('.hr-form-item__knowledgeMolds', ['知识库类别']),
                vector_model: readField('.hr-form-item__aiModelId', ['向量模型']),
                use_visual_model: (() => {
                    const visualControl = visualItem?.querySelector('input[type="checkbox"]');
                    return visualControl ? Boolean(visualControl.checked) : null;
                })(),
                knowledge_class: readField('.hr-form-item__knowledgeClassId', ['分类'])
            };
        }""")
        edit_config = normalize_edit_config(normalized_dom)
        # 以编辑页展示出来的规范化值作为后续判断依据，避免自定义下拉框
        # 的 input.value 是内部 ID 时误判向量模型。
        vector_model = edit_config["vector_model"]
        logger.info(
            "编辑页配置采集完成: fields=%s, normalizedFields=%s, knowledgeMolds=%s, vectorModel=%s",
            len(edit_form),
            len(edit_config),
            len(edit_config.get("knowledge_molds") or []),
            vector_model,
        )
        return vector_model, redact_sensitive(edit_form), redact_sensitive(edit_config)
    finally:
        _close_open_form(page)


def fetch_knowledge_detail(knowledge_id: str, token: str) -> Any:
    logger.info("开始读取知识库详情接口: knowledgeId=%s", knowledge_id)
    try:
        response = requests.get(
            DETAIL_URL,
            params={"knowledgeId": knowledge_id},
            headers={"accept": "application/json", "user-token": token},
            timeout=60,
        )
    except requests.RequestException as exc:
        logger.error("知识库详情接口请求失败: knowledgeId=%s, error=%s", knowledge_id, exc)
        raise
    if response.status_code == 403:
        fail("知识库详情接口返回 403，当前登录 token 可能无效或已过期")
    response.raise_for_status()
    payload = redact_sensitive(response.json())
    logger.info("知识库详情接口读取成功: knowledgeId=%s, httpStatus=%s", knowledge_id, response.status_code)
    return payload


PROGRESS_ID_KEYS = (
    "knowledgeProcProgressId",
    "knowledgeProcProgressID",
    "progressId",
    "progress_id",
)
FILE_ID_KEYS = ("fileId", "fileID", "file_id")
FILE_NAME_KEYS = ("fileName", "file_name", "originFileName", "originalFileName")
ENABLE_STATUS_KEYS = ("allEnable", "enableStatus", "enable_status", "enableStr", "enable_str")


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


def normalise_progress_records(payload: Any) -> list[dict[str, Any]]:
    raw_records: list[dict[str, Any]] = []
    find_progress_dicts(payload, raw_records)
    records: list[dict[str, Any]] = []
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
                # 保留规范化字段，便于下载判断和后续新增数据库功能使用。
                "knowledgeProcProgressId": progress_id,
                "fileId": first_value(raw, FILE_ID_KEYS),
                "fileName": first_value(raw, FILE_NAME_KEYS),
                "processStatus": first_value(
                    raw,
                    (
                        "procStatusStr",
                        "processStatusStr",
                        "procStatusName",
                        "processStatusName",
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
                "enableStatus": first_value(raw, ENABLE_STATUS_KEYS),
                # 接口返回的其他字段全部保留；敏感字段在写入快照时脱敏。
                "raw": redact_sensitive(raw),
            }
        )
    return records


def fetch_progress_records(knowledge_id: str, token: str) -> list[dict[str, Any]]:
    logger.info("开始读取知识库文件进度: knowledgeId=%s", knowledge_id)
    headers = {"accept": "application/json", "user-token": token}
    records: list[dict[str, Any]] = []
    seen: set[str] = set()

    for page_index in range(1, 101):
        logger.info("请求知识库文件进度分页: knowledgeId=%s, page=%s", knowledge_id, page_index)
        try:
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
        except requests.RequestException as exc:
            logger.error("知识库文件进度请求失败: knowledgeId=%s, page=%s, error=%s", knowledge_id, page_index, exc)
            raise
        if response.status_code == 403:
            fail("查询文件进度接口返回 403，当前登录 token 可能无效或已过期")
        response.raise_for_status()
        payload = response.json()
        page_records = normalise_progress_records(payload)
        records_before_page = len(records)
        for record in page_records:
            progress_id = record["knowledgeProcProgressId"]
            if progress_id and progress_id not in seen:
                seen.add(progress_id)
                records.append(record)
        logger.info(
            "知识库文件进度分页读取完成: knowledgeId=%s, page=%s, pageRecords=%s, newRecords=%s, total=%s",
            knowledge_id,
            page_index,
            len(page_records),
            len(records) - records_before_page,
            len(records),
        )

        if len(page_records) < PROGRESS_PAGE_SIZE:
            break

    logger.info("知识库文件进度读取完成: knowledgeId=%s, totalRecords=%s", knowledge_id, len(records))
    return records


STATUS_CODE_KEYS = (
    "procStatus",
    "proc_status",
    "procStatusCode",
    "processStatusCode",
    "process_status_code",
    "statusCode",
    "status_code",
    "code",
    "value",
    "dictValue",
    "dictCode",
    "itemCode",
    "statusCodeValue",
    "id",
)
STATUS_TEXT_KEYS = (
    "procStatusStr",
    "processStatusStr",
    "process_status_str",
    "statusStr",
    "status_str",
    "processStatus",
    "process_status",
    "statusName",
    "status_name",
    "name",
    "dictName",
    "dictLabel",
    "itemName",
    "statusText",
    "label",
    "text",
    "desc",
    "description",
)


def _find_status_pairs(value: Any, found: dict[str, str]) -> None:
    """从状态字典接口的不同响应包装中提取 code -> 文本映射。"""
    if isinstance(value, dict):
        code = first_value(value, STATUS_CODE_KEYS)
        text = first_value(value, STATUS_TEXT_KEYS)
        if code and text and code != text:
            found.setdefault(code, text)
        for item in value.values():
            _find_status_pairs(item, found)
    elif isinstance(value, list):
        for item in value:
            _find_status_pairs(item, found)


def fetch_progress_status_map(token: str) -> dict[str, str]:
    """读取文件处理状态字典；失败时返回空映射并继续使用记录中的状态文本。"""
    logger.info("开始读取文件处理状态字典")
    response = requests.get(
        PROC_STATUS_URL,
        headers={"accept": "application/json", "user-token": token},
        timeout=60,
    )
    if response.status_code == 403:
        fail("查询文件状态字典接口返回 403，当前登录 token 可能无效或已过期")
    response.raise_for_status()
    status_map: dict[str, str] = {}
    _find_status_pairs(response.json(), status_map)
    logger.info("文件状态字典读取完成: %s 个状态", len(status_map))
    return status_map


def _status_text(record: dict[str, Any], status_map: dict[str, str] | None = None) -> str:
    status = str(record.get("processStatus") or "").strip()
    if status:
        return status
    code = str(record.get("processStatusCode") or "").strip()
    if code and status_map:
        return str(status_map.get(code) or "").strip()
    return ""


def _is_import_success(record: dict[str, Any], status_map: dict[str, str] | None = None) -> bool:
    status = _compact_text(_status_text(record, status_map))
    if IMPORT_SUCCESS_STATUS in status:
        return True
    return (
        str(record.get("processStatusCode") or "").strip() in IMPORT_SUCCESS_STATUS_CODES
        or status in IMPORT_SUCCESS_STATUS_CODES
    )


def _is_import_failure(record: dict[str, Any], status_map: dict[str, str] | None = None) -> bool:
    status = _compact_text(_status_text(record, status_map))
    if not status:
        return False
    return any(marker in status for marker in ("失败", "错误", "异常", "拒绝", "取消"))


def _normalise_file_name(value: Any) -> str:
    name = str(value or "").strip()
    name = re.split(r"[\\/]", name)[-1]
    return name.casefold()


def list_local_xlsx_files(output_root: Path, hospital_name: str, knowledge_name: str) -> list[Path]:
    """只扫描医院/知识库目录第一层的 xlsx，自动排除 _metadata 等目录。"""
    directory = (
        output_root
        / safe_name(hospital_name, "未命名医院")
        / safe_name(knowledge_name, "未命名知识库")
    )
    if not directory.exists():
        logger.warning("本地 Excel 目录不存在: %s", directory)
        return []
    files = sorted(
        (
            path
            for path in directory.iterdir()
            if path.is_file() and path.suffix.casefold() == ".xlsx"
        ),
        key=lambda path: path.name.casefold(),
    )
    logger.info("本地 Excel 扫描完成: directory=%s, files=%s", directory, len(files))
    return files


def _extract_keyed_values(value: Any, keys: tuple[str, ...], found: list[str]) -> None:
    if isinstance(value, dict):
        for key, item in value.items():
            if str(key) in keys and item is not None and str(item).strip():
                found.append(str(item).strip())
            _extract_keyed_values(item, keys, found)
    elif isinstance(value, list):
        for item in value:
            _extract_keyed_values(item, keys, found)


def extract_progress_id(value: Any) -> str | None:
    values: list[str] = []
    _extract_keyed_values(value, PROGRESS_ID_KEYS, values)
    return values[0] if values else None


def _api_response_success(payload: Any) -> bool:
    """识别常见失败包装；没有明确失败标记时以 HTTP 2xx 为准。"""
    if not isinstance(payload, dict):
        return True
    for key in ("success", "succeed", "ok"):
        if key in payload and payload[key] is False:
            return False
    code = None
    for key in ("code", "errorCode", "error_code", "statusCode"):
        if key in payload:
            code = payload[key]
            break
    if isinstance(code, bool):
        return code
    if isinstance(code, (int, float)):
        return code in {0, 200}
    if code is not None:
        normalized = str(code).strip().casefold()
        if normalized in {"", "0", "200", "ok", "success", "succeed"}:
            return True
        if normalized in {"-1", "400", "401", "403", "404", "500", "error", "fail", "failed"}:
            return False
    return True


def _find_record_by_progress_id(
    records: list[dict[str, Any]], progress_id: str | None
) -> dict[str, Any] | None:
    if not progress_id:
        return None
    for record in records:
        if str(record.get("knowledgeProcProgressId") or "").strip() == str(progress_id).strip():
            return record
    return None


def _find_unique_new_record(
    before_records: list[dict[str, Any]],
    after_records: list[dict[str, Any]],
    file_name: str,
) -> dict[str, Any] | None:
    before_ids = {
        str(record.get("knowledgeProcProgressId") or "").strip()
        for record in before_records
        if record.get("knowledgeProcProgressId")
    }
    candidates = [
        record
        for record in after_records
        if str(record.get("knowledgeProcProgressId") or "").strip() not in before_ids
        and _normalise_file_name(record.get("fileName")) == _normalise_file_name(file_name)
    ]
    return candidates[0] if len(candidates) == 1 else None


def _target_knowledge_id_from_creation(bge_creation: Any) -> str | None:
    if not isinstance(bge_creation, dict):
        return None
    for key in ("knowledgeId", "targetKnowledgeId", "target_knowledge_id"):
        value = bge_creation.get(key)
        if value is not None and str(value).strip():
            return str(value).strip()
    return None


def resolve_bge_target_knowledge(
    page: Page,
    token: str,
    source_config: dict[str, Any],
    bge_creation: dict[str, Any] | None,
) -> tuple[str, dict[str, Any]]:
    """通过目标 ID 或目标编码找到 _bge 库，并做上传前只读校验。"""
    target_code = target_knowledge_code(source_config)
    target_id = _target_knowledge_id_from_creation(bge_creation)
    target_edit_config: dict[str, Any] | None = None
    logger.info(
        "开始解析 _bge 上传目标: sourceName=%s, targetCode=%s, creationTargetId=%s",
        source_config.get("name") or "-",
        target_code,
        target_id or "未提供",
    )

    if target_id:
        try:
            fetch_knowledge_detail(target_id, token)
            page.goto(
                f"{DETAIL_PAGE_BASE_URL}?knowledgeId={target_id}",
                wait_until="domcontentloaded",
                timeout=60_000,
            )
            _, _, target_edit_config = collect_edit_form(page)
            logger.info("已通过创建结果 knowledgeId 定位 _bge 目标库: knowledgeId=%s", target_id)
        except (requests.RequestException, VectorModelUnavailable, KnowledgeConfigurationUnavailable) as exc:
            logger.error("通过 knowledgeId 验证 _bge 目标库失败: knowledgeId=%s, error=%s", target_id, exc)
            raise KnowledgeUploadUnavailable(
                f"目标知识库 ID 无法验证: knowledgeId={target_id}，{exc}"
            ) from exc
    else:
        logger.info("创建结果未提供目标 ID，将按名称和编码搜索 _bge 目标库: code=%s", target_code)
        existing = find_knowledge_with_code(
            page,
            str(source_config.get("name") or ""),
            target_code,
        )
        if not existing:
            logger.warning("未找到 _bge 上传目标库: code=%s", target_code)
            raise KnowledgeUploadUnavailable(
                f"未找到目标知识库，且不会自动创建: code={target_code}"
            )
        target_id = str(existing["knowledgeId"])
        target_edit_config = existing["editConfig"]
        logger.info("已按编码定位 _bge 目标库: knowledgeId=%s, code=%s", target_id, target_code)

    if not target_edit_config:
        raise KnowledgeUploadUnavailable(f"未能读取目标知识库编辑配置: code={target_code}")
    actual_code = _compact_text(target_edit_config.get("code"))
    if actual_code.casefold() != _compact_text(target_code).casefold():
        logger.error("_bge 目标库编码校验失败: expected=%s, actual=%s", target_code, actual_code or "空")
        raise KnowledgeUploadUnavailable(
            f"目标知识库编码校验失败: 期望={target_code}，实际={actual_code or '空'}"
        )
    actual_model = _compact_text(target_edit_config.get("vector_model"))
    if actual_model.casefold() != BGE_VECTOR_MODEL.casefold():
        logger.error("_bge 目标库向量模型校验失败: expected=%s, actual=%s", BGE_VECTOR_MODEL, actual_model or "空")
        raise KnowledgeUploadUnavailable(
            f"目标知识库向量模型不是 {BGE_VECTOR_MODEL}: 实际={actual_model or '空'}"
        )

    upload_button = page.get_by_role("button", name="导入文件", exact=True).last
    try:
        upload_button.wait_for(state="visible", timeout=20_000)
    except PlaywrightTimeoutError as exc:
        logger.error("_bge 目标库未找到导入文件按钮: knowledgeId=%s", target_id)
        raise KnowledgeUploadUnavailable("目标知识库详情页未找到“导入文件”按钮，可能没有上传权限") from exc
    if not upload_button.is_enabled():
        logger.error("_bge 目标库导入文件按钮不可用: knowledgeId=%s", target_id)
        raise KnowledgeUploadUnavailable("目标知识库“导入文件”按钮不可用，可能没有上传权限")

    logger.info(
        "上传前目标库校验通过: knowledgeId=%s, code=%s, vectorModel=%s",
        target_id,
        target_code,
        actual_model,
    )
    return str(target_id), target_edit_config


def _close_import_dialog(page: Page) -> None:
    try:
        close_button = page.get_by_role("button", name="关闭", exact=True).last
        if close_button.count() > 0 and close_button.is_visible():
            close_button.click()
            return
    except Exception:
        pass
    _close_open_form(page)


def _open_import_dialog(page: Page) -> tuple[Any, Any]:
    logger.info("打开 _bge 文件导入窗口: page=%s", page.url)
    import_button = page.get_by_role("button", name="导入文件", exact=True).last
    try:
        import_button.wait_for(state="visible", timeout=20_000)
        if not import_button.is_enabled():
            raise KnowledgeUploadUnavailable("“导入文件”按钮不可用")
        import_button.click()
    except PlaywrightTimeoutError as exc:
        raise KnowledgeUploadUnavailable("未找到可见的“导入文件”按钮") from exc

    file_input = page.locator('input[type="file"]').last
    upload_button = page.get_by_role("button", name="上传", exact=True).last
    try:
        file_input.wait_for(state="attached", timeout=20_000)
        upload_button.wait_for(state="visible", timeout=20_000)
    except PlaywrightTimeoutError as exc:
        raise KnowledgeUploadUnavailable("打开导入窗口后未找到文件控件或“上传”按钮") from exc
    logger.info("_bge 文件导入窗口已打开")
    return file_input, upload_button


def upload_one_xlsx(
    page: Page,
    token: str,
    target_knowledge_id: str,
    local_path: Path,
    before_records: list[dict[str, Any]],
) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    """通过真实页面文件控件上传单个文件，不猜测 multipart 字段。"""
    logger.info(
        "开始上传单个 Excel: fileName=%s, targetKnowledgeId=%s, beforeRecords=%s",
        local_path.name,
        target_knowledge_id,
        len(before_records),
    )
    uploaded_at = datetime.now().astimezone().isoformat()
    result: dict[str, Any] = {
        "localPath": str(local_path.resolve()),
        "fileName": local_path.name,
        "targetKnowledgeId": target_knowledge_id,
        "uploadedAt": uploaded_at,
        "uploadStatus": "pending",
        "status": "upload_rejected",
    }
    response_payload: Any = None
    response_success = False
    response_progress_id: str | None = None
    attempted = False
    needs_reconcile = False

    try:
        file_input, upload_button = _open_import_dialog(page)
        file_input.set_input_files(str(local_path.resolve()))
        attempted = True
        logger.info("Excel 已选择，准备点击上传: fileName=%s", local_path.name)
        if not upload_button.is_enabled():
            raise KnowledgeUploadUnavailable(f"选择文件后“上传”按钮仍不可用: {local_path.name}")

        try:
            with page.expect_response(
                lambda response: (
                    UPLOAD_URL in response.url
                    and response.request.method.upper() == "POST"
                ),
                timeout=60_000,
            ) as response_info:
                upload_button.click()
            response = response_info.value
            response_success = bool(response.ok)
            try:
                response_payload = response.json()
            except Exception:
                response_payload = None
            if response_success and response_payload is not None:
                response_success = _api_response_success(response_payload)
            result["uploadHttpStatus"] = response.status
            response_progress_id = extract_progress_id(response_payload)
            if response_progress_id:
                result["knowledgeProcProgressId"] = response_progress_id
            if response_success:
                result["uploadStatus"] = "accepted"
                result["status"] = "accepted"
            else:
                result["uploadStatus"] = "rejected"
                result["status"] = "upload_rejected"
                result["error"] = f"上传接口返回失败（HTTP {response.status}）"
                logger.error("上传接口拒绝文件: fileName=%s, httpStatus=%s", local_path.name, response.status)
            needs_reconcile = response_success and not response_progress_id
            logger.info(
                "上传接口已返回: fileName=%s, httpStatus=%s, progressId=%s",
                local_path.name,
                response.status,
                response_progress_id or "未返回",
            )
        except PlaywrightTimeoutError:
            # 请求可能已到达服务器但浏览器端未收到响应；下面的文件列表对账
            # 会尝试找到唯一新增记录，找不到时绝不自动重传。
            result["uploadStatus"] = "response_timeout"
            needs_reconcile = attempted
            logger.warning("上传请求未在超时窗口内返回，将通过文件列表对账: %s", local_path.name)
        except Exception as exc:
            result["uploadStatus"] = "client_error"
            result["error"] = str(exc)
            needs_reconcile = attempted
            logger.error("触发上传失败: fileName=%s, error=%s", local_path.name, exc)
    except KnowledgeUploadUnavailable as exc:
        result["uploadStatus"] = "rejected"
        result["error"] = str(exc)
        logger.error("上传前校验失败: fileName=%s, error=%s", local_path.name, exc)
    finally:
        _close_import_dialog(page)

    after_records = before_records
    if needs_reconcile:
        try:
            after_records = fetch_progress_records(target_knowledge_id, token)
        except requests.RequestException as exc:
            result["status"] = "uploaded_untracked"
            result["uploadStatus"] = "accepted_untracked"
            result["error"] = f"上传后无法查询目标文件列表，未能关联进度 ID: {exc}"
            logger.error("上传后文件列表查询失败，禁止自动重传: fileName=%s, error=%s", local_path.name, exc)
            logger.info(
                "单个 Excel 上传处理结束: fileName=%s, status=%s, uploadStatus=%s, progressId=%s",
                local_path.name,
                result.get("status"),
                result.get("uploadStatus"),
                result.get("knowledgeProcProgressId") or "未关联",
            )
            return result, before_records

        new_record = _find_unique_new_record(before_records, after_records, local_path.name)
        if new_record:
            progress_id = str(new_record.get("knowledgeProcProgressId") or "").strip()
            if progress_id:
                result["knowledgeProcProgressId"] = progress_id
                result["uploadStatus"] = "accepted_after_reconcile"
                result["status"] = "accepted"
                result["fileId"] = new_record.get("fileId")
                result["fileStatus"] = new_record.get("processStatus")
                result["fileStatusCode"] = new_record.get("processStatusCode")
                logger.info(
                    "上传后通过文件列表唯一关联: fileName=%s, progressId=%s",
                    local_path.name,
                    progress_id,
                )
            else:
                result["status"] = "uploaded_untracked"
                result["uploadStatus"] = "accepted_untracked"
                result["error"] = "新增文件记录没有进度 ID，禁止自动重传"
                logger.warning("新增文件记录没有进度 ID，无法纳入轮询: fileName=%s", local_path.name)
        else:
            result["status"] = "uploaded_untracked"
            result["uploadStatus"] = "accepted_untracked"
            result["error"] = "上传后无法唯一关联新增文件记录，禁止自动重传"
            logger.error(
                "上传结果无法唯一关联，禁止自动重传: fileName=%s",
                local_path.name,
            )
    logger.info(
        "单个 Excel 上传处理结束: fileName=%s, status=%s, uploadStatus=%s, progressId=%s",
        local_path.name,
        result.get("status"),
        result.get("uploadStatus"),
        result.get("knowledgeProcProgressId") or "未关联",
    )
    return result, after_records


def _apply_progress_record_to_upload_result(
    result: dict[str, Any],
    record: dict[str, Any],
    status_map: dict[str, str] | None = None,
) -> None:
    result["fileId"] = record.get("fileId")
    result["knowledgeProcProgressId"] = record.get("knowledgeProcProgressId")
    result["fileStatus"] = _status_text(record, status_map) or record.get("processStatus")
    result["fileStatusCode"] = record.get("processStatusCode")
    if record.get("raw"):
        result["raw"] = record.get("raw")


def _record_is_pending_import(
    record: dict[str, Any],
    status_map: dict[str, str] | None = None,
) -> bool:
    progress_id = str(record.get("knowledgeProcProgressId") or "").strip()
    return bool(progress_id) and not _is_import_success(record, status_map) and not _is_import_failure(
        record, status_map
    )


def _aggregate_bge_upload_status(file_results: list[dict[str, Any]]) -> str:
    if not file_results:
        return "skipped_no_local_files"
    statuses = {str(item.get("status") or "") for item in file_results}
    if "uploaded_untracked" in statuses:
        return "completed_untracked"
    if "upload_rejected" in statuses or "import_failed" in statuses:
        return "completed_with_failures"
    if "timeout" in statuses:
        return "completed_with_timeouts"
    if statuses <= {"skipped_existing"}:
        return "skipped_existing"
    if statuses & {"accepted", "waiting_existing", "processing"}:
        return "pending"
    return "completed"


def poll_bge_uploads(
    target_knowledge_id: str,
    token: str,
    file_results: list[dict[str, Any]],
    status_map: dict[str, str] | None = None,
) -> list[dict[str, Any]]:
    """统一轮询一个目标知识库中本轮待处理的文件，单文件独立超时。"""
    active: dict[str, dict[str, Any]] = {}
    deadlines: dict[str, float] = {}
    poll_started = time.monotonic()
    logger.info(
        "开始轮询 _bge 文件状态: targetKnowledgeId=%s, inputFiles=%s, timeoutSeconds=%s",
        target_knowledge_id,
        len(file_results),
        UPLOAD_POLL_TIMEOUT_SECONDS,
    )

    for result in file_results:
        progress_id = str(result.get("knowledgeProcProgressId") or "").strip()
        if not progress_id or result.get("status") not in {"accepted", "waiting_existing", "processing"}:
            if result.get("status") in {"accepted", "waiting_existing", "processing"} and not progress_id:
                logger.warning("文件缺少进度 ID，无法轮询: fileName=%s", result.get("fileName") or "-")
            continue
        if progress_id in active:
            result["status"] = "uploaded_untracked"
            result["uploadStatus"] = "duplicate_progress_id"
            result["error"] = "多个文件关联到同一个进度 ID，禁止自动重传"
            logger.warning("多个文件关联到同一个进度 ID，标记为未跟踪: progressId=%s, fileName=%s", progress_id, result.get("fileName") or "-")
            continue
        active[progress_id] = result
        deadlines[progress_id] = poll_started + UPLOAD_POLL_TIMEOUT_SECONDS

    logger.info("_bge 文件状态轮询队列已建立: targetKnowledgeId=%s, activeFiles=%s", target_knowledge_id, len(active))
    poll_round = 0
    while active:
        poll_round += 1
        logger.info("刷新 _bge 文件状态: targetKnowledgeId=%s, round=%s, pendingFiles=%s", target_knowledge_id, poll_round, len(active))
        try:
            records = fetch_progress_records(target_knowledge_id, token)
        except requests.RequestException as exc:
            logger.warning("轮询目标文件状态失败，将在下一次继续: knowledgeId=%s, error=%s", target_knowledge_id, exc)
            records = []
        logger.info("_bge 文件状态刷新完成: targetKnowledgeId=%s, round=%s, records=%s", target_knowledge_id, poll_round, len(records))

        now = time.monotonic()
        for progress_id, result in list(active.items()):
            record = _find_record_by_progress_id(records, progress_id)
            if record:
                _apply_progress_record_to_upload_result(result, record, status_map)
                if _is_import_success(record, status_map):
                    result["status"] = "import_success"
                    result["uploadStatus"] = "accepted"
                    result["finishedAt"] = datetime.now().astimezone().isoformat()
                    logger.info(
                        "文件导入成功: fileName=%s, progressId=%s",
                        result.get("fileName") or "-",
                        progress_id,
                    )
                    del active[progress_id]
                    continue
                if _is_import_failure(record, status_map):
                    result["status"] = "import_failed"
                    result["uploadStatus"] = "accepted"
                    result["finishedAt"] = datetime.now().astimezone().isoformat()
                    logger.error(
                        "文件导入失败: fileName=%s, progressId=%s, status=%s",
                        result.get("fileName") or "-",
                        progress_id,
                        result.get("fileStatus") or "未知",
                    )
                    del active[progress_id]
                    continue
                result["status"] = "processing"

            if progress_id in active and now >= deadlines[progress_id]:
                result["status"] = "timeout"
                result["finishedAt"] = datetime.now().astimezone().isoformat()
                result["error"] = f"超过 {UPLOAD_POLL_TIMEOUT_SECONDS} 秒仍未完成"
                logger.error(
                    "文件导入超时: fileName=%s, progressId=%s",
                    result.get("fileName") or "-",
                    progress_id,
                )
                del active[progress_id]

        if not active:
            break

        elapsed = time.monotonic() - poll_started
        if elapsed < 30:
            interval = 2
        elif elapsed < 120:
            interval = 5
        else:
            interval = 15
        remaining = min(deadlines.values()) - time.monotonic()
        if remaining > 0:
            time.sleep(min(interval, remaining))

    logger.info("_bge 文件状态轮询结束: targetKnowledgeId=%s, rounds=%s", target_knowledge_id, poll_round)
    return file_results


def upload_bge_knowledge_base_files(
    page: Page,
    token: str,
    output_root: Path,
    hospital_name: str,
    knowledge_name: str,
    source_config: dict[str, Any],
    bge_creation: dict[str, Any] | None,
) -> dict[str, Any]:
    """准备并逐文件上传本地 Excel；返回结果，轮询由主流程统一执行。"""
    logger.info("开始准备 _bge 文件上传: knowledge=%s", knowledge_name)
    local_files = list_local_xlsx_files(output_root, hospital_name, knowledge_name)
    target_code = target_knowledge_code(source_config)
    logger.info("_bge 本地上传文件清单已确定: knowledge=%s, targetCode=%s, localFiles=%s", knowledge_name, target_code, len(local_files))
    if not local_files:
        logger.info("没有待上传的本地 Excel: knowledge=%s", knowledge_name)
        return {
            "status": "skipped_no_local_files",
            "targetCode": target_code,
            "files": [],
        }

    target_id, _ = resolve_bge_target_knowledge(page, token, source_config, bge_creation)
    existing_records = fetch_progress_records(target_id, token)
    logger.info("已读取 _bge 目标库现有文件记录: knowledge=%s, targetKnowledgeId=%s, records=%s", knowledge_name, target_id, len(existing_records))
    status_map: dict[str, str] = {}
    try:
        status_map = fetch_progress_status_map(token)
    except requests.RequestException as exc:
        logger.warning("文件状态字典读取失败，将使用文件记录中的状态: %s", exc)

    current_records = existing_records
    file_results: list[dict[str, Any]] = []
    for local_path in local_files:
        matches = [
            record
            for record in current_records
            if _normalise_file_name(record.get("fileName")) == _normalise_file_name(local_path.name)
        ]
        if matches:
            existing = matches[0]
            result: dict[str, Any] = {
                "localPath": str(local_path.resolve()),
                "fileName": local_path.name,
                "targetKnowledgeId": target_id,
                "uploadStatus": "skipped_existing",
                "observedAt": datetime.now().astimezone().isoformat(),
            }
            _apply_progress_record_to_upload_result(result, existing, status_map)
            if _record_is_pending_import(existing, status_map):
                result["status"] = "waiting_existing"
                result["uploadStatus"] = "existing_processing"
                logger.info("同名文件正在处理中，跳过上传并纳入轮询: %s", local_path.name)
            else:
                result["status"] = "skipped_existing"
                logger.info("目标知识库已存在同名文件，跳过上传: %s", local_path.name)
            file_results.append(result)
            continue

        result, after_records = upload_one_xlsx(
            page,
            token,
            target_id,
            local_path,
            current_records,
        )
        file_results.append(result)
        if after_records:
            current_records = after_records

    aggregate_status = _aggregate_bge_upload_status(file_results)
    logger.info(
        "_bge 文件上传请求处理完成: knowledge=%s, targetKnowledgeId=%s, status=%s, files=%s",
        knowledge_name,
        target_id,
        aggregate_status,
        len(file_results),
    )
    return {
        "status": aggregate_status,
        "targetKnowledgeId": target_id,
        "targetCode": target_code,
        "files": file_results,
    }


def _compact_text(value: Any) -> str:
    return re.sub(r"\s+", "", str(value or "")).strip()


def _option_text(option: Any) -> str:
    if not isinstance(option, dict):
        return str(option or "").strip()
    return str(option.get("text") or option.get("label") or option.get("value") or "").strip()


def target_knowledge_code(source_config: dict[str, Any]) -> str:
    source_code = str(source_config.get("code") or "").strip()
    if not source_code:
        raise KnowledgeConfigurationUnavailable("源知识库编码为空，无法生成目标编码")
    return f"{source_code}_bge"


def _find_form_item(drawer: Any, label: str) -> Any:
    items = drawer.locator(".hr-form-item, .hr-form__item")
    label_key = _compact_text(label)
    for index in range(items.count()):
        item = items.nth(index)
        try:
            item_text = _compact_text(item.inner_text())
        except Exception:
            continue
        if label_key in item_text:
            return item
    return None


def _drawer_field(drawer: Any, class_name: str, label: str) -> Any:
    field = drawer.locator(f".{class_name}")
    if field.count() > 0:
        return field.first
    item = _find_form_item(drawer, label)
    if item is not None:
        return item
    raise KnowledgeCreationUnavailable(f"创建表单未找到字段: {label}")


def _fill_drawer_field(drawer: Any, class_name: str, label: str, value: str) -> None:
    field = _drawer_field(drawer, class_name, label)
    control = field.locator(
        'input:not([type="radio"]):not([type="checkbox"]), textarea, [contenteditable="true"]'
    ).first
    if control.count() == 0:
        raise KnowledgeCreationUnavailable(f"创建表单字段不可输入: {label}")
    control.fill(str(value or ""))


def _control_label(control: Any) -> str:
    try:
        return str(
            control.evaluate("""(el) => {
                const text = (node) => (node?.innerText || node?.textContent || '').trim();
                const label = el.closest('label');
                if (label) return text(label);
                return text(el.parentElement) || text(el.nextElementSibling);
            }""")
            or ""
        ).strip()
    except Exception:
        return ""


def _click_visible_exact(
    page: Page,
    text: str,
    timeout_ms: int = DROPDOWN_ACTION_TIMEOUT_MS,
) -> None:
    """只在当前可见的下拉浮层内点击精确匹配项，并限制总等待时间。"""
    expected_text = _compact_text(text)
    deadline = time.monotonic() + timeout_ms / 1_000
    dropdowns = page.locator(
        "[role='listbox'], .hr-select-dropdown, .hr-select__dropdown, .hr-select__popper, "
        ".hr-popper, .hr-dropdown-menu"
    )
    option_selector = (
        "[role='option'], .hr-select-dropdown__item, .hr-select-option, "
        ".hr-dropdown__item-text, .hr-dropdown-menu__item"
    )

    while time.monotonic() < deadline:
        for dropdown_index in range(dropdowns.count()):
            dropdown = dropdowns.nth(dropdown_index)
            try:
                if not dropdown.is_visible():
                    continue
            except Exception:
                continue

            options = dropdown.locator(option_selector)
            for option_index in range(options.count()):
                option = options.nth(option_index)
                try:
                    if not option.is_visible():
                        continue
                    matches = option.evaluate(
                        r"""(el, expected) => {
                            const compact = value => String(value || '').replace(/\s+/g, '');
                            if (compact(el.innerText || el.textContent) === expected) return true;
                            return Array.from(el.querySelectorAll('*')).some(child => {
                                const style = getComputedStyle(child);
                                const rect = child.getBoundingClientRect();
                                return style.display !== 'none' && style.visibility !== 'hidden' &&
                                    Number(style.opacity || 1) !== 0 && rect.width > 0 && rect.height > 0 &&
                                    compact(child.innerText || child.textContent) === expected;
                            });
                        }""",
                        expected_text,
                    )
                    if not matches:
                        continue
                    remaining_ms = max(100, int((deadline - time.monotonic()) * 1_000))
                    option.click(timeout=remaining_ms)
                    return
                except Exception:
                    continue

        remaining_ms = int((deadline - time.monotonic()) * 1_000)
        if remaining_ms > 0:
            page.wait_for_timeout(min(100, remaining_ms))

    raise KnowledgeCreationUnavailable(
        f"下拉选项不存在或不可见（等待 {timeout_ms}ms）: {text}"
    )


def _select_drawer_option(
    drawer: Any,
    page: Page,
    class_name: str,
    label: str,
    value: str,
    fallback_value: str | None = None,
) -> str:
    field = _drawer_field(drawer, class_name, label)
    control = field.locator('input:not([type="radio"]):not([type="checkbox"])').first
    if control.count() > 0:
        control.click(timeout=DROPDOWN_ACTION_TIMEOUT_MS)
    else:
        field.click(timeout=DROPDOWN_ACTION_TIMEOUT_MS)
    selected_value = value
    requested_value = _compact_text(value)
    if fallback_value and requested_value.isdecimal():
        logger.warning(
            "创建表单分类值为纯数字，跳过原值并直接使用兜底选项: field=%s, requested=%s, fallback=%s",
            label,
            value or "未设置",
            fallback_value,
        )
        _click_visible_exact(page, fallback_value)
        selected_value = fallback_value
    else:
        try:
            _click_visible_exact(page, value)
        except KnowledgeCreationUnavailable as exc:
            if not fallback_value:
                raise
            logger.warning(
                "创建表单下拉选项不可用，使用兜底选项: field=%s, requested=%s, fallback=%s, reason=%s",
                label,
                value or "未设置",
                fallback_value,
                exc,
            )
            _click_visible_exact(page, fallback_value)
            selected_value = fallback_value
    actual_values: set[str] = set()
    try:
        if control.count() > 0:
            actual_values.add(_compact_text(control.input_value()))
    except Exception:
        pass
    try:
        field_text = _compact_text(field.inner_text())
    except Exception:
        field_text = ""
    target = _compact_text(selected_value)
    if field_text:
        actual_values.add(field_text)
    for selected in field.locator(
        '[aria-selected="true"], .hr-select__selected, .hr-select__selection,'
        ' .hr-input__inner, .hr-select__selected-item'
    ).all():
        try:
            selected_text = _compact_text(selected.inner_text() or selected.input_value())
        except Exception:
            selected_text = ""
        if selected_text:
            actual_values.add(selected_text)
    if not any(target == actual or target in actual for actual in actual_values):
        raise KnowledgeCreationUnavailable(
            f"创建表单字段选择后校验失败: {label}，期望={selected_value}，实际={' / '.join(actual_values) or '未知'}"
        )
    return selected_value


def _set_radio_option(drawer: Any, label: str, option: dict[str, Any]) -> None:
    item = _find_form_item(drawer, "私有库类型")
    if item is None:
        raise KnowledgeCreationUnavailable("创建表单未找到字段: 私有库类型")
    radios = item.locator('input[type="radio"]')
    expected_value = str(option.get("value") or "").strip()
    expected_text = _option_text(option)
    selected = None
    for index in range(radios.count()):
        radio = radios.nth(index)
        radio_value = str(radio.get_attribute("value") or "").strip()
        radio_text = _compact_text(_control_label(radio))
        if (expected_value and radio_value == expected_value) or (
            expected_text and _compact_text(expected_text) in radio_text
        ):
            selected = radio
            break
    if selected is None:
        raise KnowledgeCreationUnavailable(f"创建表单无法匹配私有库类型: {expected_text or expected_value}")
    if not selected.is_checked():
        selected.evaluate("(el) => el.click()")
    if not selected.is_checked():
        raise KnowledgeCreationUnavailable(f"私有库类型选择后校验失败: {expected_text or expected_value}")


def _set_checkbox_options(drawer: Any, source_options: list[dict[str, Any]]) -> None:
    item = _find_form_item(drawer, "知识库类别")
    if item is None:
        item = drawer.locator(".hr-form-item__knowledgeMolds").first
    if item is None or item.count() == 0:
        raise KnowledgeCreationUnavailable("创建表单未找到字段: 知识库类别")

    source_by_value = {
        str(option.get("value") or "").strip(): option
        for option in source_options
        if str(option.get("value") or "").strip()
    }
    source_by_text = {
        _compact_text(_option_text(option)): option
        for option in source_options
        if _compact_text(_option_text(option))
    }
    matched_source: set[int] = set()
    checkboxes = item.locator('input[type="checkbox"]')
    if checkboxes.count() == 0:
        raise KnowledgeCreationUnavailable("创建表单知识库类别没有可操作的复选框")

    for index in range(checkboxes.count()):
        checkbox = checkboxes.nth(index)
        value = str(checkbox.get_attribute("value") or "").strip()
        text = _compact_text(_control_label(checkbox))
        source = source_by_value.get(value) or source_by_text.get(text)
        desired = bool(source and source.get("checked"))
        if source:
            matched_source.add(id(source))
        if checkbox.is_checked() != desired:
            checkbox.evaluate("(el) => el.click()")
        if checkbox.is_checked() != desired:
            raise KnowledgeCreationUnavailable(f"知识库类别选择后校验失败: {text or value}")

    missing_checked = [
        _option_text(option)
        for option in source_options
        if option.get("checked") and id(option) not in matched_source
    ]
    if missing_checked:
        raise KnowledgeCreationUnavailable(
            "创建表单无法匹配源知识库类别: " + ", ".join(missing_checked)
        )


def _set_visual_model(drawer: Any, enabled: bool) -> None:
    field = drawer.locator(".hr-form-item__useVLModel")
    if field.count() == 0:
        field = _find_form_item(drawer, "启动视觉模型")
    if field is None or field.count() == 0:
        raise KnowledgeCreationUnavailable("创建表单未找到字段: 启动视觉模型")
    checkbox = field.locator('input[type="checkbox"]').first
    if checkbox.count() == 0:
        raise KnowledgeCreationUnavailable("创建表单启动视觉模型字段不可操作")
    if checkbox.is_checked() != enabled:
        checkbox.evaluate("(el) => el.click()")
    if checkbox.is_checked() != enabled:
        raise KnowledgeCreationUnavailable("启动视觉模型选择后校验失败")


def _open_create_drawer(page: Page) -> Any:
    logger.info("打开知识库创建入口: listUrl=%s", KNOWLEDGE_LIST_URL)
    page.goto(KNOWLEDGE_LIST_URL, wait_until="domcontentloaded", timeout=60_000)
    list_search_box = page.get_by_placeholder(SEARCH_PLACEHOLDER, exact=True)
    new_button = page.get_by_role("button", name="新增知识库", exact=True)
    try:
        # 先确认列表页的异步内容已经挂载，避免只看到外壳按钮就开始操作。
        list_search_box.wait_for(state="visible", timeout=20_000)
        new_button.wait_for(state="visible", timeout=20_000)
    except PlaywrightTimeoutError:
        raise KnowledgeCreationUnavailable("知识库列表页未找到右上角“新增知识库”按钮") from None

    new_button.first.click()
    logger.info("已点击“新增知识库”主按钮，等待创建菜单")

    # 下拉菜单由前端异步渲染，不能在 click 后立即 count()。按优先级使用
    # 语义定位和已知 class 定位，并让 Locator 自己等待元素真正出现。
    menu_candidates = [
        page.get_by_role("menuitem", name="新增知识库", exact=True),
        page.locator(".hr-dropdown__item-text").filter(has_text="新增知识库"),
        page.locator(".hr-dropdown-menu__item").filter(has_text="新增知识库"),
        page.locator("[class*='dropdown'] [class*='item']").filter(has_text="新增知识库"),
    ]
    menu_error: Exception | None = None
    for candidate in menu_candidates:
        try:
            candidate.first.wait_for(state="visible", timeout=5_000)
            candidate.first.click()
            logger.info("已选择“新增知识库”菜单项")
            break
        except Exception as exc:
            menu_error = exc
    else:
        raise KnowledgeCreationUnavailable(
            "点击“新增知识库”后未找到可见的二级菜单项"
        ) from menu_error

    drawer = page.locator(".hr-drawer.hr-drawer--open").last
    try:
        drawer.wait_for(state="visible", timeout=20_000)
    except PlaywrightTimeoutError:
        raise KnowledgeCreationUnavailable("未打开知识库创建抽屉表单") from None
    logger.info("知识库创建抽屉已打开")
    return drawer


def build_expected_bge_config(source_config: dict[str, Any], target_code: str) -> dict[str, Any]:
    expected = json.loads(json.dumps(source_config, ensure_ascii=False))
    expected["code"] = target_code
    expected["vector_model"] = BGE_VECTOR_MODEL
    return expected


def _field_matches(expected: Any, actual: Any) -> bool:
    expected_values = set()
    actual_values = set()
    for item, target in ((expected, expected_values), (actual, actual_values)):
        if isinstance(item, dict):
            for key in ("value", "text", "label"):
                value = _compact_text(item.get(key))
                if value:
                    target.add(value)
        else:
            value = _compact_text(item)
            if value:
                target.add(value)
    return bool(expected_values & actual_values)


def _checked_option_keys(options: Any) -> set[str]:
    result: set[str] = set()
    if not isinstance(options, list):
        return result
    for option in options:
        if not isinstance(option, dict) or not option.get("checked"):
            continue
        key = _compact_text(option.get("value")) or _compact_text(_option_text(option))
        if key:
            result.add(key)
    return result


def compare_knowledge_configs(expected: dict[str, Any], actual: dict[str, Any]) -> list[str]:
    mismatches: list[str] = []
    for key in ("name", "code", "description", "vector_model"):
        if _compact_text(expected.get(key)) != _compact_text(actual.get(key)):
            mismatches.append(key)
    for key in ("scope", "private_type", "knowledge_class"):
        if not _field_matches(expected.get(key), actual.get(key)):
            mismatches.append(key)
    if bool(expected.get("use_visual_model")) != bool(actual.get("use_visual_model")):
        mismatches.append("use_visual_model")
    if _checked_option_keys(expected.get("knowledge_molds")) != _checked_option_keys(
        actual.get("knowledge_molds")
    ):
        mismatches.append("knowledge_molds")
    return mismatches


def find_knowledge_with_code(page: Page, knowledge_name: str, target_code: str) -> dict[str, Any] | None:
    """搜索同名知识库并逐个读取编辑页编码，避免只按名称误判。"""
    logger.info("开始按名称和编码查找知识库: name=%s, code=%s", knowledge_name, target_code)
    try:
        matched_count, _ = search_knowledge(page, knowledge_name)
    except KnowledgeNotFound:
        logger.info("未找到同名知识库，编码查找结束: name=%s, code=%s", knowledge_name, target_code)
        return None
    logger.info("发现同名知识库候选: name=%s, count=%s", knowledge_name, matched_count)
    for index in range(matched_count):
        logger.info("读取同名候选的编辑配置: name=%s, candidate=%s/%s", knowledge_name, index + 1, matched_count)
        try:
            _, knowledge_id = search_knowledge(page, knowledge_name, index)
            _, edit_form, edit_config = collect_edit_form(page)
        except KnowledgeNotFound:
            continue
        except (VectorModelUnavailable, KnowledgeConfigurationUnavailable) as exc:
            # 无法确认某个同名结果的编码时，不能安全判断目标编码是否已存在；
            # 创建流程应整体跳过，避免误创建重复知识库。
            raise KnowledgeCreationUnavailable(
                f"无法确认同名知识库第 {index + 1} 项的编码: {exc}"
            ) from exc
        if _compact_text(edit_config.get("code")).casefold() == _compact_text(target_code).casefold():
            logger.info("找到目标知识库: name=%s, code=%s, knowledgeId=%s", knowledge_name, target_code, knowledge_id)
            return {
                "knowledgeId": knowledge_id,
                "editForm": edit_form,
                "editConfig": edit_config,
            }
    logger.info("同名候选中未找到目标编码: name=%s, code=%s", knowledge_name, target_code)
    return None


def verify_created_knowledge(
    page: Page,
    expected_config: dict[str, Any],
    actual_knowledge_class: str | None = None,
) -> dict[str, Any]:
    logger.info("开始复核新建知识库: name=%s, code=%s", expected_config.get("name") or "-", expected_config.get("code") or "-")
    existing = find_knowledge_with_code(
        page,
        str(expected_config.get("name") or ""),
        str(expected_config.get("code") or ""),
    )
    if not existing:
        logger.error("创建后复核未找到目标知识库: code=%s", expected_config.get("code") or "-")
        raise KnowledgeCreationUnavailable(
            f"创建后重新搜索未找到目标知识库编码: {expected_config.get('code')}"
        )
    verification_config = json.loads(json.dumps(expected_config, ensure_ascii=False))
    requested_knowledge_class = _edit_field_text(expected_config.get("knowledge_class"))
    selected_knowledge_class = _compact_text(actual_knowledge_class)
    fallback_used = bool(
        selected_knowledge_class
        and _compact_text(selected_knowledge_class) != _compact_text(requested_knowledge_class)
    )
    if selected_knowledge_class:
        # 创建表单可能只能显示兜底分类的文本，详情页返回的 value/text
        # 也可能与源库的历史值不同；复核应以本次实际选择为准。
        verification_config["knowledge_class"] = {
            "value": selected_knowledge_class,
            "text": selected_knowledge_class,
        }
    mismatches = compare_knowledge_configs(verification_config, existing["editConfig"])
    result = {
        "status": "verified" if not mismatches else "verification_failed",
        "knowledgeId": existing["knowledgeId"],
        "targetCode": expected_config.get("code"),
        "requestedKnowledgeClass": requested_knowledge_class or None,
        "actualKnowledgeClass": selected_knowledge_class or None,
        "fallbackUsed": fallback_used,
        "mismatches": mismatches,
        "actualEditConfig": existing["editConfig"],
    }
    if mismatches:
        logger.error("创建后知识库字段复核失败: code=%s, mismatches=%s", expected_config.get("code") or "-", mismatches)
        raise KnowledgeCreationUnavailable(
            "创建后字段复核不一致: " + ", ".join(mismatches)
        )
    logger.info(
        "创建后知识库字段复核成功: code=%s, knowledgeId=%s, requestedKnowledgeClass=%s, actualKnowledgeClass=%s, fallbackUsed=%s",
        expected_config.get("code") or "-",
        existing["knowledgeId"],
        requested_knowledge_class or "未设置",
        selected_knowledge_class or "未返回",
        fallback_used,
    )
    return result


def create_bge_knowledge_base(
    page: Page,
    source_config: dict[str, Any],
) -> dict[str, Any]:
    target_code = target_knowledge_code(source_config)
    logger.info("开始创建 _bge 知识库: sourceName=%s, targetCode=%s", source_config.get("name") or "-", target_code)
    existing = find_knowledge_with_code(page, str(source_config.get("name") or ""), target_code)
    if existing:
        logger.warning("目标编码已存在，跳过创建且不修改已有知识库: %s", target_code)
        return {
            "status": "skipped_existing",
            "targetCode": target_code,
            "knowledgeId": existing["knowledgeId"],
        }

    expected_config = build_expected_bge_config(source_config, target_code)
    logger.info("_bge 创建配置已生成: targetCode=%s, vectorModel=%s, knowledgeMolds=%s", target_code, BGE_VECTOR_MODEL, len(expected_config.get("knowledge_molds") or []))
    drawer = _open_create_drawer(page)
    try:
        _fill_drawer_field(drawer, "hr-form-item__knowledgeTypeName", "知识库名称", expected_config["name"])
        _fill_drawer_field(drawer, "hr-form-item__knowledgeTypeCode", "知识库编码", expected_config["code"])
        _fill_drawer_field(
            drawer,
            "hr-form-item__knowledgeTypeDesc",
            "知识库说明",
            expected_config.get("description") or "",
        )
        _select_drawer_option(
            drawer,
            page,
            "hr-form-item__knowledgeScap",
            "使用范围",
            _edit_field_text(expected_config["scope"]),
        )
        _set_radio_option(drawer, "私有库类型", expected_config["private_type"])
        _set_checkbox_options(drawer, expected_config["knowledge_molds"])
        _select_drawer_option(
            drawer,
            page,
            "hr-form-item__aiModelId",
            "向量模型",
            BGE_VECTOR_MODEL,
        )
        _set_visual_model(drawer, bool(expected_config.get("use_visual_model")))
        selected_knowledge_class = _select_drawer_option(
            drawer,
            page,
            "hr-form-item__knowledgeClassId",
            "分类",
            _edit_field_text(expected_config["knowledge_class"]),
            fallback_value="智能导诊",
        )
        logger.info("_bge 创建表单字段已填充，开始检查确认按钮: targetCode=%s", target_code)

        confirm = drawer.locator("button.hr-drawer__confirm").last
        if confirm.count() == 0 or not confirm.is_visible():
            raise KnowledgeCreationUnavailable("创建抽屉未找到可见的“确认”按钮")
        if confirm.is_disabled():
            raise KnowledgeCreationUnavailable("创建表单校验未通过，“确认”按钮处于禁用状态")
        logger.info("_bge 创建表单校验通过，点击确认: targetCode=%s", target_code)
        confirm.click()
        try:
            drawer.wait_for(state="hidden", timeout=30_000)
        except PlaywrightTimeoutError:
            raise KnowledgeCreationUnavailable("点击确认后创建抽屉未关闭，无法确认提交结果") from None
    except KnowledgeCreationUnavailable:
        _close_open_form(page)
        raise
    except Exception as exc:
        _close_open_form(page)
        raise KnowledgeCreationUnavailable(f"创建表单操作失败: {exc}") from exc

    logger.info("知识库创建请求已提交: name=%s, code=%s", expected_config["name"], target_code)
    return verify_created_knowledge(page, expected_config, selected_knowledge_class)


def save_knowledge_snapshot(
    output_root: Path,
    hospital_name: str,
    knowledge_name: str,
    knowledge_id: str,
    detail_page_url: str,
    detail_api: Any,
    edit_form: list[dict[str, Any]],
    progress_records: list[dict[str, Any]],
    vector_model: str,
    edit_config: dict[str, Any] | None = None,
    bge_creation: dict[str, Any] | None = None,
    bge_upload: dict[str, Any] | None = None,
) -> Path:
    """保存知识库详情、编辑页字段和文件列表，供后续新增数据库功能复用。"""
    metadata_dir = output_root / safe_name(hospital_name, "未命名医院") / safe_name(knowledge_name, "未命名知识库") / "_metadata"
    logger.info(
        "开始保存知识库快照: knowledge=%s, knowledgeId=%s, progressRecords=%s, bgeCreation=%s, bgeUpload=%s",
        knowledge_name,
        knowledge_id,
        len(progress_records),
        bge_creation.get("status") if isinstance(bge_creation, dict) else "未提供",
        bge_upload.get("status") if isinstance(bge_upload, dict) else "未提供",
    )
    try:
        metadata_dir.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        logger.error("知识库快照目录创建失败: directory=%s, error=%s", metadata_dir, exc)
        raise
    files: list[dict[str, Any]] = []
    for record in progress_records:
        files.append(
            {
                "fileName": record.get("fileName"),
                "fileId": record.get("fileId"),
                "knowledgeProcProgressId": record.get("knowledgeProcProgressId"),
                "fileStatus": record.get("processStatus"),
                "fileStatusCode": record.get("processStatusCode"),
                "enableStatus": record.get("enableStatus"),
                "raw": record.get("raw", {}),
            }
        )
    snapshot = {
        "capturedAt": datetime.now().astimezone().isoformat(),
        "hospitalName": hospital_name,
        "knowledgeName": knowledge_name,
        "knowledgeId": knowledge_id,
        "detailPageUrl": detail_page_url,
        "vectorModel": vector_model,
        "detailApi": redact_sensitive(detail_api),
        "editForm": redact_sensitive(edit_form),
        "editConfig": redact_sensitive(edit_config or {}),
        "files": redact_sensitive(files),
    }
    snapshot_path = metadata_dir / "knowledge_snapshot.json"
    if bge_creation is not None:
        snapshot["bgeCreation"] = redact_sensitive(bge_creation)
    elif snapshot_path.exists():
        try:
            with snapshot_path.open("r", encoding="utf-8") as previous_file:
                previous_snapshot = json.load(previous_file)
            if isinstance(previous_snapshot, dict) and previous_snapshot.get("bgeCreation") is not None:
                snapshot["bgeCreation"] = redact_sensitive(previous_snapshot["bgeCreation"])
        except (OSError, json.JSONDecodeError):
            pass
    if bge_upload is not None:
        snapshot["bgeUpload"] = redact_sensitive(bge_upload)
    elif snapshot_path.exists():
        try:
            with snapshot_path.open("r", encoding="utf-8") as previous_file:
                previous_snapshot = json.load(previous_file)
            if isinstance(previous_snapshot, dict) and previous_snapshot.get("bgeUpload") is not None:
                snapshot["bgeUpload"] = redact_sensitive(previous_snapshot["bgeUpload"])
        except (OSError, json.JSONDecodeError):
            pass
    try:
        with snapshot_path.open("w", encoding="utf-8") as snapshot_file:
            json.dump(snapshot, snapshot_file, ensure_ascii=False, indent=2)
            snapshot_file.write("\n")
    except (OSError, TypeError, ValueError):
        logger.exception("知识库快照写入失败: path=%s", snapshot_path)
        raise
    logger.info(
        "知识库快照写入成功: path=%s, files=%s, hasBgeCreation=%s, hasBgeUpload=%s",
        snapshot_path,
        len(files),
        "bgeCreation" in snapshot,
        "bgeUpload" in snapshot,
    )
    return snapshot_path


def get_knowledge_snapshot_path(
    output_root: Path,
    hospital_name: str,
    knowledge_name: str,
) -> Path:
    return (
        output_root
        / safe_name(hospital_name, "未命名医院")
        / safe_name(knowledge_name, "未命名知识库")
        / "_metadata"
        / "knowledge_snapshot.json"
    )


def load_knowledge_snapshot(
    output_root: Path,
    hospital_name: str,
    knowledge_name: str,
) -> tuple[Path, dict[str, Any]]:
    snapshot_path = get_knowledge_snapshot_path(output_root, hospital_name, knowledge_name)
    if not snapshot_path.exists():
        raise KnowledgeConfigurationUnavailable(f"未找到已有知识库快照: {snapshot_path}")
    try:
        with snapshot_path.open("r", encoding="utf-8") as snapshot_file:
            snapshot = json.load(snapshot_file)
    except (OSError, json.JSONDecodeError) as exc:
        raise KnowledgeConfigurationUnavailable(f"读取知识库快照失败: {snapshot_path}: {exc}") from exc
    if not isinstance(snapshot, dict):
        raise KnowledgeConfigurationUnavailable(f"知识库快照格式无效: {snapshot_path}")
    edit_config = snapshot.get("editConfig")
    if not isinstance(edit_config, dict) or not str(edit_config.get("code") or "").strip():
        raise KnowledgeConfigurationUnavailable(f"知识库快照缺少有效 editConfig/code: {snapshot_path}")
    return snapshot_path, snapshot


def update_bge_snapshot(
    snapshot_path: Path,
    bge_creation: dict[str, Any] | None = None,
    bge_upload: dict[str, Any] | None = None,
) -> None:
    try:
        with snapshot_path.open("r", encoding="utf-8") as snapshot_file:
            snapshot = json.load(snapshot_file)
        if not isinstance(snapshot, dict):
            raise ValueError("快照根节点不是对象")
        if bge_creation is not None:
            snapshot["bgeCreation"] = redact_sensitive(bge_creation)
        if bge_upload is not None:
            snapshot["bgeUpload"] = redact_sensitive(bge_upload)
        with snapshot_path.open("w", encoding="utf-8") as snapshot_file:
            json.dump(snapshot, snapshot_file, ensure_ascii=False, indent=2)
            snapshot_file.write("\n")
    except (OSError, TypeError, ValueError, json.JSONDecodeError):
        logger.exception("_bge 后处理快照更新失败: path=%s", snapshot_path)
        raise
    logger.info(
        "_bge 后处理快照已更新: path=%s, creation=%s, upload=%s",
        snapshot_path,
        bge_creation.get("status") if bge_creation else "未更新",
        bge_upload.get("status") if bge_upload else "未更新",
    )


def run_bge_only(
    page: Page,
    token: str,
    output_root: Path,
    hospital_name: str,
    knowledge_names: list[str],
    create_bge_enabled: bool,
    upload_bge_enabled: bool,
) -> None:
    """复用已有快照和本地 Excel，仅执行 _bge 创建、上传和状态轮询。"""
    logger.info("进入仅后处理模式：跳过源知识条目读取和 Excel 导出")
    pending_uploads: list[dict[str, Any]] = []
    summary: list[tuple[str, str, str, str]] = []
    processed_count = 0

    for knowledge_name in knowledge_names:
        logger.info("开始仅后处理知识库: %s", knowledge_name)
        try:
            snapshot_path, snapshot = load_knowledge_snapshot(
                output_root,
                hospital_name,
                knowledge_name,
            )
        except KnowledgeConfigurationUnavailable as exc:
            logger.warning("跳过仅后处理知识库: %s", exc)
            summary.append((knowledge_name, "快照不可用", "未执行", "未执行"))
            continue

        edit_config = snapshot["editConfig"]
        vector_model = str(snapshot.get("vectorModel") or edit_config.get("vector_model") or "").strip()
        target_code = target_knowledge_code(edit_config)
        local_files = list_local_xlsx_files(output_root, hospital_name, knowledge_name)
        if vector_model.casefold() != TARGET_VECTOR_MODEL.casefold():
            bge_creation = {
                "status": "skipped_vector_model",
                "targetCode": target_code,
                "sourceVectorModel": vector_model,
                "targetVectorModel": BGE_VECTOR_MODEL,
            }
            bge_upload = {
                "status": "skipped_vector_model",
                "targetCode": target_code,
                "files": [],
            }
            update_bge_snapshot(snapshot_path, bge_creation, bge_upload)
            summary.append((knowledge_name, "跳过", bge_creation["status"], bge_upload["status"]))
            continue

        if not local_files:
            bge_creation = {
                "status": "skipped_no_local_files",
                "targetCode": target_code,
                "reason": "没有可供后处理的本地 Excel",
            }
            bge_upload = {
                "status": "skipped_no_local_files",
                "targetCode": target_code,
                "files": [],
            }
            update_bge_snapshot(snapshot_path, bge_creation, bge_upload)
            logger.warning("跳过仅后处理知识库: 没有本地 Excel: %s", knowledge_name)
            summary.append((knowledge_name, "跳过", bge_creation["status"], bge_upload["status"]))
            continue

        processed_count += 1
        if not create_bge_enabled:
            bge_creation = {
                "status": "disabled",
                "targetCode": target_code,
                "reason": "创建开关未开启",
            }
            logger.info("未创建 _bge 知识库: 创建开关未开启")
        else:
            try:
                bge_creation = create_bge_knowledge_base(page, edit_config)
            except KnowledgeCreationUnavailable as exc:
                bge_creation = {
                    "status": "failed",
                    "targetCode": target_code,
                    "reason": str(exc),
                }
                logger.error("仅后处理创建 _bge 失败，继续处理后续知识库: %s", exc)

        bge_upload: dict[str, Any] | None = None
        if upload_bge_enabled:
            try:
                bge_upload = upload_bge_knowledge_base_files(
                    page,
                    token,
                    output_root,
                    hospital_name,
                    knowledge_name,
                    edit_config,
                    bge_creation,
                )
                logger.info(
                    "仅后处理上传准备完成: knowledge=%s, status=%s, files=%s",
                    knowledge_name,
                    bge_upload.get("status"),
                    len(bge_upload.get("files", [])),
                )
            except (KnowledgeUploadUnavailable, requests.RequestException) as exc:
                bge_upload = {
                    "status": "skipped_target_unavailable",
                    "targetCode": target_code,
                    "files": [],
                    "error": str(exc),
                }
                logger.error("跳过仅后处理上传: knowledge=%s, error=%s", knowledge_name, exc)
        else:
            bge_upload = {
                "status": "disabled",
                "targetCode": target_code,
                "files": [],
                "reason": "上传开关未开启",
            }

        update_bge_snapshot(snapshot_path, bge_creation, bge_upload)
        if any(
            item.get("status") in {"accepted", "waiting_existing", "processing"}
            for item in bge_upload.get("files", [])
        ):
            pending_uploads.append(
                {
                    "targetKnowledgeId": bge_upload.get("targetKnowledgeId"),
                    "bgeUpload": bge_upload,
                    "snapshotPath": snapshot_path,
                    "knowledgeName": knowledge_name,
                    "bgeCreation": bge_creation,
                }
            )
        summary.append(
            (
                knowledge_name,
                "完成",
                bge_creation.get("status", "未知"),
                bge_upload.get("status", "未知"),
            )
        )

    if pending_uploads:
        try:
            upload_status_map = fetch_progress_status_map(token)
        except requests.RequestException as exc:
            upload_status_map = {}
            logger.warning("仅后处理统一轮询前读取状态字典失败，将使用文件记录中的状态: %s", exc)

        logger.info("开始统一轮询仅后处理上传文件状态: %s 个知识库", len(pending_uploads))
        for job in pending_uploads:
            target_id = str(job.get("targetKnowledgeId") or "").strip()
            if not target_id:
                logger.warning("跳过状态轮询：上传结果缺少 targetKnowledgeId: %s", job["knowledgeName"])
                continue
            final_files = poll_bge_uploads(
                target_id,
                token,
                job["bgeUpload"].get("files", []),
                upload_status_map,
            )
            job["bgeUpload"]["files"] = final_files
            job["bgeUpload"]["status"] = _aggregate_bge_upload_status(final_files)
            update_bge_snapshot(
                job["snapshotPath"],
                job["bgeCreation"],
                job["bgeUpload"],
            )
            logger.info(
                "仅后处理上传状态轮询完成: knowledge=%s, status=%s",
                job["knowledgeName"],
                job["bgeUpload"]["status"],
            )
            for index, item in enumerate(summary):
                if item[0] == job["knowledgeName"]:
                    summary[index] = (
                        item[0],
                        item[1],
                        item[2],
                        job["bgeUpload"]["status"],
                    )
                    break

    if processed_count == 0:
        logger.warning("仅后处理没有找到包含本地 Excel 的可处理知识库")
    logger.info("仅后处理运行汇总:")
    for knowledge_name, status, creation_status, upload_status in summary:
        logger.info(
            "- %s: %s；_bge 创建=%s；_bge 上传=%s",
            knowledge_name,
            status,
            creation_status,
            upload_status,
        )


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
    parser.add_argument("--log-dir", help=f"日志目录，默认: {DEFAULT_LOG_ROOT}")
    parser.add_argument(
        "--create-bge",
        action="store_true",
        help="满足全部条件时创建 _bge 知识库；默认关闭，不会提交创建",
    )
    parser.add_argument(
        "--upload-bge",
        action="store_true",
        help="将本地 Excel 上传到已存在的 _bge 知识库并轮询导入状态；默认关闭",
    )
    stage_group = parser.add_mutually_exclusive_group()
    stage_group.add_argument(
        "--export-only",
        action="store_true",
        help="仅执行源知识库读取和 Excel 导出，跳过 _bge 创建与上传",
    )
    stage_group.add_argument(
        "--create-only",
        action="store_true",
        help="跳过导出，仅复用已有快照和本地 Excel 创建 _bge 知识库，不上传文件",
    )
    stage_group.add_argument(
        "--upload-only",
        action="store_true",
        help="跳过导出和创建，仅使用已有快照及 Excel 上传文件并轮询导入状态",
    )
    stage_group.add_argument(
        "--bge-only",
        action="store_true",
        help="跳过源知识库导出，复用已有快照和本地 Excel；创建和上传仍由配置或对应开关控制",
    )
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


def is_create_bge_enabled(args: argparse.Namespace, config: dict[str, Any]) -> bool:
    """读取创建开关；命令行显式开启优先于配置文件。"""
    if getattr(args, "create_bge", False):
        return True
    value = config.get("createBgeKnowledgeBases", False)
    if isinstance(value, bool):
        return value
    return str(value).strip().casefold() in {"1", "true", "yes", "y", "on"}


def is_upload_bge_enabled(args: argparse.Namespace, config: dict[str, Any]) -> bool:
    """读取上传开关；命令行显式开启优先于配置文件。"""
    if getattr(args, "upload_bge", False):
        return True
    value = config.get("uploadBgeKnowledgeBases", False)
    if isinstance(value, bool):
        return value
    return str(value).strip().casefold() in {"1", "true", "yes", "y", "on"}


def main() -> int:
    args = parse_args()
    global logger
    log_dir = Path(args.log_dir).expanduser().resolve() if args.log_dir else DEFAULT_LOG_ROOT
    logger, log_path = configure_logging(log_dir, "knowledge_query", "export_knowledge")
    logger.info("开始自动批量导出")
    config_path = Path(args.config).expanduser().resolve()
    config = load_config(config_path)
    username, password = get_credentials(args, config)
    knowledge_names = get_knowledge_names(args, config)
    if args.export_only:
        create_bge_enabled = False
        upload_bge_enabled = False
        run_postprocessing_only = False
        logger.info("运行模式: 仅导出，忽略 _bge 创建和上传开关")
    elif args.create_only:
        create_bge_enabled = True
        upload_bge_enabled = False
        run_postprocessing_only = True
        logger.info("运行模式: 仅创建 _bge，跳过导出和文件上传")
    elif args.upload_only:
        create_bge_enabled = False
        upload_bge_enabled = True
        run_postprocessing_only = True
        logger.info("运行模式: 仅上传文件，跳过导出和 _bge 创建")
    else:
        create_bge_enabled = is_create_bge_enabled(args, config)
        upload_bge_enabled = is_upload_bge_enabled(args, config)
        run_postprocessing_only = args.bge_only
    hospital_name = str(config_value(config, "hospitalName", "hospital_name", default="demo"))
    output_root_value = config_value(config, "outputRoot", "output_root")
    output_root = Path(output_root_value).expanduser() if output_root_value else DEFAULT_OUTPUT_ROOT
    if not output_root.is_absolute():
        output_root = (SCRIPT_DIR / output_root).resolve()
    logger.info(
        "运行配置读取完成: config=%s, hospital=%s, knowledgeCount=%s, outputRoot=%s",
        config_path,
        hospital_name,
        len(knowledge_names),
        output_root,
    )
    logger.info("登录凭据已准备（账号、密码和 token 不写入日志）")
    logger.info(
        "_bge 知识库创建开关: %s",
        "已开启（仅满足全部条件时提交）" if create_bge_enabled else "未开启，不会点击确认",
    )
    logger.info(
        "_bge 文件上传开关: %s",
        "已开启（会向线上目标库传输本地 Excel）" if upload_bge_enabled else "未开启，不会选择文件或调用上传流程",
    )

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
            logger.info("登录流程完成，开始处理知识库列表: count=%s", len(knowledge_names))
            if run_postprocessing_only:
                token = token_holder.get("value")
                if not token:
                    fail("登录成功但未捕获到页面请求中的 user-token，无法执行 _bge 后处理")
                run_bge_only(
                    page,
                    token,
                    output_root,
                    hospital_name,
                    knowledge_names,
                    create_bge_enabled,
                    upload_bge_enabled,
                )
                logger.info("仅后处理运行结束，日志文件: %s", log_path)
                return 0
            total_exported = 0
            summary: list[tuple[str, str, int, int, int]] = []
            pending_uploads: list[dict[str, Any]] = []

            for knowledge_name in knowledge_names:
                logger.info("开始处理知识库: %s", knowledge_name)
                try:
                    matched_count, knowledge_id = search_knowledge(page, knowledge_name)
                except KnowledgeNotFound as exc:
                    logger.warning("跳过知识库: %s", exc)
                    summary.append((knowledge_name, "未找到", 0, 0, 0))
                    continue

                try:
                    vector_model, edit_form, edit_config = collect_edit_form(page)
                except (VectorModelUnavailable, KnowledgeConfigurationUnavailable) as exc:
                    logger.warning(
                        "跳过知识库: 无法安全读取编辑页配置（不会使用默认值）: %s",
                        exc,
                    )
                    summary.append((knowledge_name, "编辑页配置读取失败", 0, 0, 0))
                    continue

                token = token_holder.get("value")
                if not token:
                    fail("登录成功但未捕获到页面请求中的 user-token，无法调用文件进度接口")

                try:
                    detail_api = fetch_knowledge_detail(knowledge_id, token)
                except requests.RequestException as exc:
                    logger.warning("知识库详情接口获取失败，将保留编辑页和文件列表信息: %s", exc)
                    detail_api = {"error": str(exc)}

                try:
                    progress_records = fetch_progress_records(knowledge_id, token)
                except requests.RequestException as exc:
                    logger.warning("跳过知识库: 获取文件进度失败 (%s)", exc)
                    snapshot_path = save_knowledge_snapshot(
                        output_root,
                        hospital_name,
                        knowledge_name,
                        knowledge_id,
                        page.url,
                        detail_api,
                        edit_form,
                        [],
                        vector_model,
                        edit_config=edit_config,
                        bge_creation={
                            "status": "skipped_progress_fetch_failed",
                            "reason": str(exc),
                        },
                    )
                    logger.info("知识库快照已保存: %s | 编辑字段 %s 项 | 文件记录 0 条", snapshot_path, len(edit_form))
                    summary.append((knowledge_name, "获取文件进度失败", 0, 0, 0))
                    continue

                source_detail_page_url = page.url
                snapshot_path = save_knowledge_snapshot(
                    output_root,
                    hospital_name,
                    knowledge_name,
                    knowledge_id,
                    page.url,
                    detail_api,
                    edit_form,
                    progress_records,
                    vector_model,
                    edit_config=edit_config,
                )
                logger.info(
                    "知识库快照已保存: %s | 编辑字段 %s 项 | 文件记录 %s 条",
                    snapshot_path,
                    len(edit_form),
                    len(progress_records),
                )
                logger.info("向量模型: %s", vector_model)

                # 下载条件顺序：先判断向量模型，再判断文件状态。
                if vector_model.casefold() != TARGET_VECTOR_MODEL.casefold():
                    bge_creation = {
                        "status": "skipped_vector_model",
                        "sourceVectorModel": vector_model,
                        "targetVectorModel": BGE_VECTOR_MODEL,
                    }
                    logger.warning(
                        f"跳过知识库下载: 向量模型为 {vector_model}，"
                        f"仅 {TARGET_VECTOR_MODEL} 允许下载"
                    )
                    snapshot_path = save_knowledge_snapshot(
                        output_root,
                        hospital_name,
                        knowledge_name,
                        knowledge_id,
                        source_detail_page_url,
                        detail_api,
                        edit_form,
                        progress_records,
                        vector_model,
                        edit_config=edit_config,
                        bge_creation=bge_creation,
                    )
                    summary.append((knowledge_name, "向量模型不匹配，已跳过", len(progress_records), 0, 0))
                    continue
                logger.info("向量模型符合条件，继续判断文件状态: %s", TARGET_VECTOR_MODEL)

                successful_records: list[dict[str, Any]] = []
                for record in progress_records:
                    process_status = (record.get("processStatus") or "").strip()
                    if _is_import_success(record):
                        successful_records.append(record)
                    else:
                        logger.info(
                            f"跳过未导入成功文件: "
                            f"progressId={record.get('knowledgeProcProgressId') or '-'} "
                            f"| fileName={record.get('fileName') or '-'} "
                            f"| 状态={process_status or '未知'}"
                        )

                logger.info("进入详情页并获取文件进度 ID 成功")
                logger.info("详情页: %s", page.url)
                logger.info("知识库: %s", knowledge_name)
                logger.info("匹配到的同名文本数量: %s", matched_count)
                logger.info("knowledgeId: %s", knowledge_id)
                logger.info("文件进度记录数: %s", len(progress_records))
                logger.info("导入成功记录数: %s", len(successful_records))

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
                        logger.error("导出失败: progressId=%s (%s)", progress_id, exc)
                        continue
                    if output_path:
                        exported_for_knowledge += 1
                        total_exported += 1
                        logger.info(
                            f"导出完成: {output_path}"
                            f" | fileName={record.get('fileName') or '-'}"
                            f" | fileId={record.get('fileId') or '-'}"
                            f" | 数据条数={row_count}"
                        )
                    else:
                        logger.warning("跳过空文件: progressId=%s", progress_id)

                if not create_bge_enabled:
                    bge_creation = {
                        "status": "disabled",
                        "targetCode": target_knowledge_code(edit_config),
                        "reason": "配置项 createBgeKnowledgeBases 未开启，且未传 --create-bge",
                    }
                    logger.info("未创建 _bge 知识库: 创建开关未开启")
                elif not successful_records:
                    bge_creation = {
                        "status": "skipped_no_import_success",
                        "targetCode": target_knowledge_code(edit_config),
                        "reason": f"没有文件状态为“{IMPORT_SUCCESS_STATUS}”",
                    }
                    logger.warning("未创建 _bge 知识库: 没有导入成功的文件")
                elif exported_for_knowledge < 1:
                    bge_creation = {
                        "status": "skipped_no_export",
                        "targetCode": target_knowledge_code(edit_config),
                        "reason": "没有成功导出任何文件",
                    }
                    logger.warning("未创建 _bge 知识库: 没有成功导出任何文件")
                else:
                    try:
                        bge_creation = create_bge_knowledge_base(page, edit_config)
                    except KnowledgeCreationUnavailable as exc:
                        bge_creation = {
                            "status": "failed",
                            "targetCode": target_knowledge_code(edit_config),
                            "reason": str(exc),
                        }
                        logger.error("创建 _bge 知识库失败，继续处理后续知识库: %s", exc)

                bge_upload: dict[str, Any] | None = None
                if upload_bge_enabled:
                    try:
                        bge_upload = upload_bge_knowledge_base_files(
                            page,
                            token,
                            output_root,
                            hospital_name,
                            knowledge_name,
                            edit_config,
                            bge_creation,
                        )
                        if any(
                            item.get("status") in {"accepted", "waiting_existing", "processing"}
                            for item in bge_upload.get("files", [])
                        ):
                            pending_uploads.append(
                                {
                                    "targetKnowledgeId": bge_upload.get("targetKnowledgeId"),
                                    "bgeUpload": bge_upload,
                                    "outputRoot": output_root,
                                    "hospitalName": hospital_name,
                                    "knowledgeName": knowledge_name,
                                    "knowledgeId": knowledge_id,
                                    "detailPageUrl": source_detail_page_url,
                                    "detailApi": detail_api,
                                    "editForm": edit_form,
                                    "progressRecords": progress_records,
                                    "vectorModel": vector_model,
                                    "editConfig": edit_config,
                                    "bgeCreation": bge_creation,
                                }
                            )
                        logger.info(
                            "_bge 文件上传准备完成: knowledge=%s, status=%s, files=%s",
                            knowledge_name,
                            bge_upload.get("status"),
                            len(bge_upload.get("files", [])),
                        )
                    except (KnowledgeUploadUnavailable, requests.RequestException) as exc:
                        bge_upload = {
                            "status": "skipped_target_unavailable",
                            "targetCode": target_knowledge_code(edit_config),
                            "files": [],
                            "error": str(exc),
                        }
                        logger.error(
                            "跳过 _bge 文件上传: knowledge=%s, error=%s",
                            knowledge_name,
                            exc,
                        )

                snapshot_path = save_knowledge_snapshot(
                    output_root,
                    hospital_name,
                    knowledge_name,
                    knowledge_id,
                    source_detail_page_url,
                    detail_api,
                    edit_form,
                    progress_records,
                    vector_model,
                    edit_config=edit_config,
                    bge_creation=bge_creation,
                    bge_upload=bge_upload,
                )
                logger.info(
                    "知识库快照已更新: %s | _bge 创建状态=%s | _bge 上传状态=%s",
                    snapshot_path,
                    bge_creation["status"],
                    bge_upload.get("status") if bge_upload else "未执行",
                )

                summary.append(
                    (
                        knowledge_name,
                        "完成",
                        len(progress_records),
                        len(successful_records),
                        exported_for_knowledge,
                    )
                )

            if pending_uploads:
                try:
                    upload_status_map = fetch_progress_status_map(token_holder.get("value", ""))
                except requests.RequestException as exc:
                    upload_status_map = {}
                    logger.warning("统一轮询前读取状态字典失败，将使用文件记录中的状态: %s", exc)

                logger.info("开始统一轮询全部 _bge 文件状态: %s 个知识库", len(pending_uploads))
                for job in pending_uploads:
                    target_id = str(job.get("targetKnowledgeId") or "").strip()
                    if not target_id:
                        continue
                    bge_upload = poll_bge_uploads(
                        target_id,
                        token_holder.get("value", ""),
                        job["bgeUpload"].get("files", []),
                        upload_status_map,
                    )
                    job["bgeUpload"]["files"] = bge_upload
                    job["bgeUpload"]["status"] = _aggregate_bge_upload_status(bge_upload)
                    final_snapshot_path = save_knowledge_snapshot(
                        job["outputRoot"],
                        job["hospitalName"],
                        job["knowledgeName"],
                        job["knowledgeId"],
                        job["detailPageUrl"],
                        job["detailApi"],
                        job["editForm"],
                        job["progressRecords"],
                        job["vectorModel"],
                        edit_config=job["editConfig"],
                        bge_creation=job["bgeCreation"],
                        bge_upload=job["bgeUpload"],
                    )
                    logger.info(
                        "_bge 文件状态轮询完成: knowledge=%s, status=%s, snapshot=%s",
                        job["knowledgeName"],
                        job["bgeUpload"]["status"],
                        final_snapshot_path,
                    )

            logger.info("本次运行汇总:")
            for knowledge_name, status, progress_count, successful_count, exported_count in summary:
                logger.info(
                    f"- {knowledge_name}: {status}；"
                    f"文件记录 {progress_count} 条，导入成功 {successful_count} 条，"
                    f"成功导出 {exported_count} 个文件"
                )
            if total_exported == 0:
                fail("四个知识库均没有导出任何数据")
            logger.info("本次运行结束，成功导出 %s 个文件，日志文件: %s", total_exported, log_path)
            return 0
        except Exception:
            logger.exception("本次运行发生未预期异常")
            raise
        finally:
            browser.close()
            logger.info("浏览器已关闭")


if __name__ == "__main__":
    raise SystemExit(main())
