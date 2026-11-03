# -*- coding: utf-8 -*-
"""AI 助手：调用 OpenAI 兼容接口（配置见 data\\ai_config.json）。
设计原则：AI 只做「理解和提取」，返回内容必须通过确定性校验器才能进入业务流程；
未配置密钥或调用失败时，所有调用方必须能无损退回纯手工流程。"""
import datetime
import hashlib
import json
import os
import re
import urllib.error
import urllib.request

APP = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(APP, "data", "ai_config.json")

_CACHE = {"mtime": None, "config": {}}


def load_config():
    """读取配置（带缓存，改配置文件立即生效）。"""
    try:
        mtime = os.path.getmtime(CONFIG_PATH)
    except OSError:
        return {}
    if _CACHE["mtime"] != mtime:
        try:
            with open(CONFIG_PATH, encoding="utf-8-sig") as fp:
                config = json.load(fp)
        except (OSError, ValueError):
            config = {}
        _CACHE["mtime"] = mtime
        _CACHE["config"] = config if isinstance(config, dict) else {}
    return _CACHE["config"]


def is_configured():
    config = load_config()
    return bool(config.get("enabled") and config.get("base_url")
                and config.get("api_key") and config.get("model"))


def status():
    """给前端的状态：不含密钥。"""
    config = load_config()
    return {
        "configured": is_configured(),
        "model": config.get("model", ""),
        "base_url": config.get("base_url", ""),
        "note": "" if is_configured() else
                "未启用 AI（检查 data\\ai_config.json），相关按钮不可用",
    }


class AiError(RuntimeError):
    pass


# 轻量进度状态：前端轮询判断 AI 进行到哪一步（进行中/卡住/失败）
PROGRESS = {"stage": "", "detail": "", "at": "", "done": True, "error": ""}


def set_progress(stage, detail="", done=False, error=""):
    PROGRESS.update({"stage": stage, "detail": str(detail or ""),
                     "at": datetime.datetime.now().isoformat(timespec="seconds"),
                     "done": bool(done), "error": str(error or "")})


def _extract_json(text):
    """从回复中抠出第一个完整 JSON 对象（容忍代码围栏与前后废话）。"""
    if not text:
        raise AiError("AI 返回为空")
    text = re.sub(r"```(?:json)?|```", "", str(text))
    start = text.find("{")
    if start < 0:
        raise AiError("AI 回复中未找到 JSON")
    depth = 0
    in_string = False
    escape = False
    for index in range(start, len(text)):
        char = text[index]
        if in_string:
            if escape:
                escape = False
            elif char == "\\":
                escape = True
            elif char == '"':
                in_string = False
            continue
        if char == '"':
            in_string = True
        elif char == "{":
            depth += 1
        elif char == "}":
            depth -= 1
            if depth == 0:
                candidate = text[start:index + 1]
                try:
                    return json.loads(candidate)
                except ValueError as exc:
                    raise AiError("AI 返回的 JSON 无法解析：%s" % exc) from exc
    raise AiError("AI 返回的 JSON 不完整")


def chat(messages, config=None):
    """调用 /chat/completions，返回助手文本。失败抛 AiError。"""
    config = config or load_config()
    if not config.get("base_url") or not config.get("api_key"):
        raise AiError("AI 未配置（data\\ai_config.json）")
    url = config["base_url"].rstrip("/") + "/chat/completions"
    payload = {
        "model": config.get("model", ""),
        "messages": messages,
        "temperature": float(config.get("temperature", 0.1)),
        "stream": False,
    }
    request = urllib.request.Request(
        url,
        data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
        headers={"Content-Type": "application/json",
                 "Authorization": "Bearer %s" % config["api_key"]},
        method="POST")
    timeout = float(config.get("timeout_seconds") or 60)
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            body = json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detail = ""
        try:
            detail = exc.read().decode("utf-8", "ignore")[:300]
        except Exception:
            pass
        raise AiError("AI 接口 HTTP %s%s" % (exc.code, "：" + detail if detail else "")) from exc
    except (urllib.error.URLError, OSError, ValueError) as exc:
        raise AiError("AI 接口连接失败：%s" % exc) from exc
    try:
        return body["choices"][0]["message"]["content"]
    except (KeyError, IndexError, TypeError) as exc:
        raise AiError("AI 接口返回结构异常：%s" % str(body)[:200]) from exc


def extract_json(system_prompt, user_prompt, validate=None, max_retries=2):
    """带校验重试的提取：AI 输出非法 JSON 或未通过 validate 时，把错误喂回去重问。
    返回 (data, meta)；meta 记录模型、时间与响应哈希，用于留痕。"""
    config = load_config()
    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": user_prompt},
    ]
    last_error = None
    for attempt in range(max_retries + 1):
        set_progress("调用模型", "第 %d/%d 次尝试（单次最长 %d 秒）…"
                     % (attempt + 1, max_retries + 1,
                        int(float(config.get("timeout_seconds") or 60))))
        text = chat(messages, config)
        set_progress("模型已返回", "正在解析与校验数据…")
        try:
            data = _extract_json(text)
            if validate is not None:
                validate(data)
            set_progress("完成", "提取成功（第 %d 次尝试）" % (attempt + 1), done=True)
            return data, {
                "model": config.get("model", ""),
                "at": datetime.datetime.now().isoformat(timespec="seconds"),
                "response_sha256": hashlib.sha256(
                    text.encode("utf-8")).hexdigest()[:16],
                "attempts": attempt + 1,
            }
        except (AiError, ValueError) as exc:
            last_error = exc
            set_progress("上次输出不合格", "正在把问题喂回模型重试…")
            messages.append({"role": "assistant", "content": text[:2000]})
            messages.append({"role": "user", "content":
                             "你的输出存在问题：%s\n请重新只输出一个符合要求的 JSON 对象，"
                             "不要解释、不要代码块标记。" % exc})
    set_progress("提取失败", str(last_error), done=True, error=str(last_error))
    raise AiError("AI 提取重试 %d 次仍未通过校验：%s" % (max_retries + 1, last_error))
