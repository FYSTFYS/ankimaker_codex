#!/usr/bin/env python3
import argparse
import base64
from concurrent.futures import ThreadPoolExecutor, as_completed
import html
import json
import os
import re
import sys
import time
import unicodedata
import urllib.error
import urllib.request
import urllib.parse
import datetime as dt
from pathlib import Path
from typing import Any

from run_utils import StepLogger, write_text, write_xlsx


ANKI_CONNECT_URL = os.getenv("ANKI_CONNECT_URL", "http://127.0.0.1:8765")
DEFAULT_LLM_PROVIDER = os.getenv("LLM_PROVIDER", "doubao").lower()
OPENAI_API_URL = os.getenv("OPENAI_API_URL", "https://api.openai.com/v1/responses")
DOUBAO_API_URL = os.getenv("DOUBAO_API_URL", "https://ark.cn-beijing.volces.com/api/v3/responses")
OPENAI_IMAGE_API_URL = os.getenv("OPENAI_IMAGE_API_URL", "https://api.openai.com/v1/images/generations")
PIXABAY_API_URL = os.getenv("PIXABAY_API_URL", "https://pixabay.com/api/")
DEFAULT_OPENAI_MODEL = os.getenv("OPENAI_MODEL", "gpt-4.1-mini")
DEFAULT_DOUBAO_MODEL = os.getenv("DOUBAO_MODEL", "doubao-seed-2-0-lite-260428")
DEFAULT_IMAGE_MODEL = os.getenv("OPENAI_IMAGE_MODEL", "gpt-image-1")
DEFAULT_DECK = os.getenv("ANKI_DECK", "English::AI Words")
DEFAULT_NOTE_TYPE = os.getenv("ANKI_NOTE_TYPE", "AI English Word")
DEFAULT_CONFIG = "anki_config.json"
DEFAULT_INTERNET_IMAGE_LIMIT = 3
DEFAULT_WORD_PROMPT = """
我要你帮我记忆单词，我给出一个单词：{word}

规则：
- 如果这个单词不是真实的英文单词，只输出：非英文单词
- 如果是真实英文单词，只输出 JSON，不要解释，不要代码块，不要多余文字
- JSON 必须包含这些字段：
  - word：单词原形
  - phonetic：音标
  - meanings_image：常见释义及图片记忆，至少 2 项
  - related_words：相关词汇数组，每项必须是对象，格式为 {"word":"...","pos":"...","meaning":"..."}
  - collocations：地道用法/固定搭配
  - memory_method：记忆方法
- 内容要短、清楚、适合记忆
""".strip()

FIELDS = [
    "Word",
    "Phonetic",
    "MemoryImage",
    "MeaningsImage",
    "RelatedWords",
    "Collocations",
    "MemoryMethod",
]


class NonEnglishWordError(RuntimeError):
    pass


CARD_CSS = """
.card {
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Arial, sans-serif;
  font-size: 18px;
  line-height: 1.55;
  color: #1f2933;
  background: #fbfaf7;
  text-align: left;
}
.word {
  font-size: 34px;
  font-weight: 750;
  color: #172554;
  margin-bottom: 8px;
}
.phonetic {
  color: #52616b;
  font-size: 20px;
  margin-bottom: 14px;
}
.section {
  border-top: 1px solid #d7d2c8;
  padding-top: 10px;
  margin-top: 12px;
}
.label {
  color: #8a4b0f;
  font-weight: 700;
  margin-bottom: 4px;
}
ul {
  margin-top: 4px;
  padding-left: 22px;
}
""".strip()


FRONT_TEMPLATE = """
<div class="word">{{Word}}</div>
<div class="phonetic">{{Phonetic}}</div>
""".strip()


BACK_TEMPLATE = """
{{FrontSide}}

<div class="section">
  {{MemoryImage}}
</div>

<div class="section">
  <div class="label">常见含义和图片记忆</div>
  <div>{{MeaningsImage}}</div>
</div>

<div class="section">
  <div class="label">相关词扩展</div>
  <div>{{RelatedWords}}</div>
</div>

<div class="section">
  <div class="label">常见搭配</div>
  <div>{{Collocations}}</div>
</div>

<div class="section">
  <div class="label">记忆方法</div>
  <div>{{MemoryMethod}}</div>
</div>
""".strip()


WORD_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "word": {"type": "string"},
        "phonetic": {"type": "string"},
        "meanings_image": {
            "type": "array",
            "items": {
                "type": "object",
                "additionalProperties": False,
                "properties": {
                    "meaning": {"type": "string"},
                    "image_memory": {"type": "string"},
                },
                "required": ["meaning", "image_memory"],
            },
        },
        "related_words": {
            "type": "array",
            "items": {"type": "string"},
        },
        "collocations": {
            "type": "array",
            "items": {
                "type": "object",
                "additionalProperties": False,
                "properties": {
                    "phrase": {"type": "string"},
                    "translation": {"type": "string"},
                    "example": {"type": "string"},
                },
                "required": ["phrase", "translation", "example"],
            },
        },
        "memory_method": {"type": "string"},
    },
    "required": [
        "word",
        "phonetic",
        "meanings_image",
        "related_words",
        "collocations",
        "memory_method",
    ],
}


def post_json(url: str, payload: dict[str, Any], headers: dict[str, str] | None = None) -> dict[str, Any]:
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    request_headers = {"Content-Type": "application/json", **(headers or {})}
    request = urllib.request.Request(url, data=body, headers=request_headers)
    try:
        with urllib.request.urlopen(request, timeout=90) as response:
            return json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"HTTP {exc.code} from {url}: {detail}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"Cannot reach {url}: {exc.reason}") from exc


def invoke_anki(action: str, params: dict[str, Any] | None = None) -> Any:
    response = post_json(
        ANKI_CONNECT_URL,
        {"action": action, "version": 6, "params": params or {}},
    )
    if response.get("error"):
        raise RuntimeError(f"AnkiConnect {action} failed: {response['error']}")
    return response.get("result")


def ensure_deck(deck: str) -> None:
    invoke_anki("createDeck", {"deck": deck})


def ensure_note_type(note_type: str) -> None:
    models = invoke_anki("modelNames")
    if note_type not in models:
        invoke_anki(
            "createModel",
            {
                "modelName": note_type,
                "inOrderFields": FIELDS,
                "css": CARD_CSS,
                "cardTemplates": [
                    {
                        "Name": "Recognition",
                        "Front": FRONT_TEMPLATE,
                        "Back": BACK_TEMPLATE,
                    }
                ],
            },
        )
        return

    existing_fields = invoke_anki("modelFieldNames", {"modelName": note_type})
    for field in FIELDS:
        if field not in existing_fields:
            invoke_anki("modelFieldAdd", {"modelName": note_type, "fieldName": field})


def load_config(path: str) -> dict[str, Any]:
    if not os.path.exists(path):
        return {}
    with open(path, "r", encoding="utf-8") as handle:
        return json.load(handle)


def resolve_relative_path(path_value: str | Path, base_dir: Path) -> Path:
    path = Path(path_value)
    if path.is_absolute():
        return path
    return base_dir / path


def resolve_output_path(
    configured_value: Any,
    default_value: str,
    base_dir: Path,
    stem: str | None,
    derived_name: str,
) -> Path:
    value = str(configured_value or default_value)
    if stem and value == default_value:
        value = derived_name
    return resolve_relative_path(value, base_dir)


def resolve_llm_settings(args: argparse.Namespace, config: dict[str, Any]) -> dict[str, str]:
    llm_config = config.get("llm", {})
    provider = (args.provider or llm_config.get("provider") or DEFAULT_LLM_PROVIDER).lower()
    if provider not in {"openai", "doubao"}:
        raise RuntimeError(f"Unsupported LLM provider: {provider}")

    default_model = DEFAULT_DOUBAO_MODEL if provider == "doubao" else DEFAULT_OPENAI_MODEL
    api_url = (
        args.api_url
        or llm_config.get("api_url")
        or (DOUBAO_API_URL if provider == "doubao" else OPENAI_API_URL)
    )
    model = args.model or llm_config.get("model") or default_model
    return {
        "provider": provider,
        "api_url": api_url,
        "model": model,
        "api_key": resolve_llm_api_key(provider, llm_config),
        "prompt_template": str(llm_config.get("prompt_template") or DEFAULT_WORD_PROMPT),
    }


def resolve_llm_api_key(provider: str, llm_config: dict[str, Any]) -> str:
    configured_key = llm_config.get("api_key")
    if isinstance(configured_key, str) and configured_key.strip():
        return configured_key.strip()

    env_names = ["LLM_API_KEY"]
    if provider == "doubao":
        env_names.append("ARK_API_KEY")
    else:
        env_names.append("OPENAI_API_KEY")

    for env_name in env_names:
        api_key = os.getenv(env_name)
        if api_key:
            return api_key

    expected = "ARK_API_KEY" if provider == "doubao" else "OPENAI_API_KEY"
    raise RuntimeError(f"{expected} is not set.")


def resolve_image_settings(config: dict[str, Any]) -> dict[str, Any]:
    image_config = config.get("image", {})
    configured_limit = image_config.get("per_word_limit", DEFAULT_INTERNET_IMAGE_LIMIT)
    configured_per_page = image_config.get("per_page", DEFAULT_INTERNET_IMAGE_LIMIT)
    try:
        limit = int(configured_limit)
    except (TypeError, ValueError):
        limit = DEFAULT_INTERNET_IMAGE_LIMIT
    try:
        per_page = int(configured_per_page)
    except (TypeError, ValueError):
        per_page = DEFAULT_INTERNET_IMAGE_LIMIT

    return {
        "source": str(image_config.get("source") or "generated"),
        "provider": str(image_config.get("provider") or "pixabay").lower(),
        "api_key": str(image_config.get("api_key") or os.getenv("PIXABAY_API_KEY") or "").strip(),
        "per_word_limit": max(1, min(limit, DEFAULT_INTERNET_IMAGE_LIMIT)),
        "per_page": max(3, per_page),
    }


def read_words(args: argparse.Namespace) -> list[str]:
    if args.entries_json:
        return []

    chunks: list[str] = []
    if args.file:
        with open(args.file, "r", encoding="utf-8") as handle:
            chunks.append(handle.read())
    if args.words:
        chunks.append(" ".join(args.words))
    if not chunks and not sys.stdin.isatty():
        chunks.append(sys.stdin.read())

    words: list[str] = []
    seen: set[str] = set()
    for raw in re.split(r"[\s,，;；]+", "\n".join(chunks)):
        word = raw.strip()
        if not word:
            continue
        key = word.lower()
        if key not in seen:
            seen.add(key)
            words.append(word)
    return words


def extract_response_text(response: dict[str, Any]) -> str:
    if isinstance(response.get("output_text"), str):
        return response["output_text"]

    texts: list[str] = []
    for item in response.get("output", []):
        for content in item.get("content", []):
            text = content.get("text")
            if isinstance(text, str):
                texts.append(text)
    return "\n".join(texts).strip()


def coerce_list(value: Any) -> list[Any]:
    if isinstance(value, list):
        return value
    if value is None:
        return []
    return [value]


def normalize_meanings_image(value: Any) -> list[dict[str, str]]:
    items = coerce_list(value)
    normalized: list[dict[str, str]] = []
    for item in items:
        if isinstance(item, dict):
            meaning = item.get("meaning") or item.get("translation") or item.get("text") or ""
            image_memory = item.get("image_memory") or item.get("image") or item.get("memory") or meaning
            normalized.append(
                {
                    "meaning": str(meaning).strip(),
                    "image_memory": str(image_memory).strip(),
                }
            )
        else:
            text = str(item).strip()
            if text:
                normalized.append(
                    {
                        "meaning": text,
                        "image_memory": text,
                    }
                )
    if len(normalized) == 1:
        normalized.append(
            {
                "meaning": normalized[0]["meaning"],
                "image_memory": normalized[0]["image_memory"],
            }
        )
    return normalized


def normalize_entry(entry: dict[str, Any], word: str) -> dict[str, Any]:
    normalized = dict(entry)
    normalized["word"] = str(normalized.get("word") or word).strip() or word
    normalized["phonetic"] = str(normalized.get("phonetic") or "").strip()
    normalized["meanings_image"] = normalize_meanings_image(normalized.get("meanings_image"))
    normalized["related_words"] = [
        {
            "word": str(item.get("word") or item.get("term") or item.get("related_word") or "").strip(),
            "pos": str(item.get("pos") or item.get("part_of_speech") or "").strip(),
            "meaning": str(item.get("meaning") or item.get("translation") or "").strip(),
        }
        if isinstance(item, dict)
        else {
            "word": str(item).strip(),
            "pos": "",
            "meaning": "",
        }
        for item in coerce_list(normalized.get("related_words"))
        if str(item).strip()
    ]
    normalized["collocations"] = [
        (
            {
                "phrase": str(item.get("phrase") or item.get("collocation") or item.get("text") or "").strip(),
                "translation": str(item.get("translation") or item.get("meaning") or "").strip(),
                "example": str(item.get("example") or item.get("sentence") or "").strip(),
            }
            if isinstance(item, dict)
            else {
                "phrase": str(item).strip(),
                "translation": str(item).strip(),
                "example": str(item).strip(),
            }
        )
        for item in coerce_list(normalized.get("collocations"))
        if str(item).strip()
    ]
    normalized["memory_method"] = str(normalized.get("memory_method") or "").strip()
    image_files = normalized.get("image_files")
    if isinstance(image_files, list):
        normalized["image_files"] = [str(item).strip() for item in image_files if str(item).strip()]
    elif normalized.get("image_file"):
        normalized["image_files"] = [str(normalized["image_file"]).strip()]
    else:
        normalized["image_files"] = []
    return normalized


def generate_word_entry(word: str, model: str, api_url: str, api_key: str, prompt_template: str) -> dict[str, Any]:
    prompt = prompt_template.replace("{word}", word)

    response = post_json(
        api_url,
        {
            "model": model,
            "input": prompt,
        },
        {"Authorization": f"Bearer {api_key}"},
    )
    text = extract_response_text(response)
    if not text:
        raise RuntimeError(f"OpenAI returned no text for {word!r}.")
    if text.strip() == "非英文单词":
        raise NonEnglishWordError(word)
    return normalize_entry(json.loads(text), word)


def fetch_json(url: str) -> dict[str, Any]:
    request = urllib.request.Request(
        url,
        headers={
            "User-Agent": "ankimaker/0.1 (Anki vocabulary image lookup)",
        },
    )
    try:
        with urllib.request.urlopen(request, timeout=30) as response:
            return json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"HTTP {exc.code} from {url}: {detail}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"Cannot reach {url}: {exc.reason}") from exc


def fetch_bytes(url: str) -> bytes:
    request = urllib.request.Request(
        url,
        headers={
            "User-Agent": "ankimaker/0.1 (Anki vocabulary image lookup)",
        },
    )
    with urllib.request.urlopen(request, timeout=60) as response:
        return response.read()


def load_entries(path: str) -> list[dict[str, Any]]:
    with open(path, "r", encoding="utf-8") as handle:
        data = json.load(handle)
    if isinstance(data, dict):
        data = data.get("entries")
    if not isinstance(data, list):
        raise RuntimeError("--entries-json must contain a JSON array or an object with an entries array.")
    for entry in data:
        missing = [
            key
            for key in [
                "word",
                "phonetic",
                "meanings_image",
                "related_words",
                "collocations",
                "memory_method",
            ]
            if key not in entry
        ]
        if missing:
            raise RuntimeError(f"Entry is missing required keys {missing}: {entry!r}")
    return data


def slugify(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", value).encode("ascii", "ignore").decode("ascii")
    slug = re.sub(r"[^a-zA-Z0-9]+", "-", normalized).strip("-").lower()
    return slug or "word"


def build_image_prompt(entry: dict[str, Any]) -> str:
    scenes = "; ".join(item["image_memory"] for item in entry["meanings_image"][:3])
    meanings = "; ".join(item["meaning"] for item in entry["meanings_image"][:3])
    return f"""
Create one vivid mnemonic illustration for the English vocabulary word "{entry["word"]}".

Meanings to suggest visually: {meanings}
Mnemonic scene: {scenes}

Style: clear educational illustration, memorable, colorful, simple composition, suitable for an Anki flashcard.
Important: no text, no letters, no captions, no watermark.
""".strip()


def generate_memory_image(entry: dict[str, Any], image_model: str, image_size: str, image_quality: str) -> bytes:
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        raise RuntimeError("OPENAI_API_KEY is not set.")

    response = post_json(
        OPENAI_IMAGE_API_URL,
        {
            "model": image_model,
            "prompt": build_image_prompt(entry),
            "size": image_size,
            "quality": image_quality,
            "n": 1,
        },
        {"Authorization": f"Bearer {api_key}"},
    )

    data = response.get("data") or []
    if not data or not data[0].get("b64_json"):
        raise RuntimeError(f"OpenAI returned no image data for {entry['word']!r}.")
    return base64.b64decode(data[0]["b64_json"])


def find_pixabay_images(word: str, api_key: str, limit: int, per_page: int) -> list[dict[str, str]]:
    if not api_key:
        raise RuntimeError("Pixabay API key is not configured.")

    query = urllib.parse.urlencode(
        {
            "key": api_key,
            "q": word,
            "image_type": "photo",
            "safesearch": "true",
            "order": "popular",
            "per_page": per_page,
        }
    )
    data = fetch_json(f"{PIXABAY_API_URL}?{query}")
    hits = data.get("hits") or []
    results: list[dict[str, str]] = []
    for hit in hits[:limit]:
        url = hit.get("largeImageURL") or hit.get("webformatURL") or hit.get("previewURL")
        if not url:
            continue
        user = str(hit.get("user") or "").strip()
        credit = "Pixabay"
        if user:
            credit += f", by {user}"
        results.append(
            {
                "url": str(url),
                "source_page": str(hit.get("pageURL") or url),
                "credit": credit,
            }
        )
    return results


def image_extension_from_url(url: str, default: str = ".jpg") -> str:
    suffix = Path(urllib.parse.urlparse(url).path).suffix.lower()
    if suffix in {".jpg", ".jpeg", ".png", ".webp", ".gif"}:
        return suffix
    return default


def download_internet_images(
    entry: dict[str, Any],
    image_dir: str,
    provider: str,
    api_key: str,
    limit: int,
    per_page: int,
) -> tuple[list[str], list[str]]:
    if provider != "pixabay":
        raise RuntimeError(f"Unsupported internet image provider: {provider}")

    results = find_pixabay_images(entry["word"], api_key, limit, per_page)
    if not results:
        raise RuntimeError(f"No Pixabay image found for {entry['word']!r}.")

    paths: list[str] = []
    source_urls: list[str] = []
    credits: list[str] = []
    for index, result in enumerate(results, start=1):
        image_bytes = fetch_bytes(result["url"])
        path = save_image_file(
            image_dir,
            entry["word"],
            image_bytes,
            index=index,
            extension=image_extension_from_url(result["url"]),
        )
        paths.append(path)
        source_urls.append(result["source_page"] or result["url"])
        credits.append(result["credit"])

    entry["image_files"] = paths
    entry["image_file"] = paths[0]
    entry["image_source_urls"] = source_urls
    entry["image_source_url"] = "\n".join(source_urls)
    entry["image_credits"] = credits
    entry["image_credit"] = credits[0] if credits else ""
    return paths, source_urls


def save_image_file(
    image_dir: str,
    word: str,
    image_bytes: bytes,
    *,
    index: int | None = None,
    extension: str = ".png",
) -> str:
    os.makedirs(image_dir, exist_ok=True)
    suffix = f"-{index}" if index is not None else ""
    filename = f"aiword-{slugify(word)}{suffix}{extension}"
    path = os.path.join(image_dir, filename)
    with open(path, "wb") as handle:
        handle.write(image_bytes)
    return path


def store_anki_media(filename: str, image_bytes: bytes) -> str:
    invoke_anki(
        "storeMediaFile",
        {
            "filename": filename,
            "data": base64.b64encode(image_bytes).decode("ascii"),
        },
    )
    return filename


def store_anki_media_from_path(path: str) -> str:
    with open(path, "rb") as handle:
        image_bytes = handle.read()
    filename = os.path.basename(path)
    return store_anki_media(filename, image_bytes)


def store_anki_media_from_paths(paths: list[str]) -> list[str]:
    return [store_anki_media_from_path(path) for path in paths]


def get_image_paths(entry: dict[str, Any]) -> list[str]:
    image_files = entry.get("image_files")
    if isinstance(image_files, list):
        paths = [str(path).strip() for path in image_files if str(path).strip()]
        if paths:
            return paths
    image_file = str(entry.get("image_file") or "").strip()
    return [image_file] if image_file else []


def html_list(items: list[Any]) -> str:
    rows = []
    for item in items:
        if isinstance(item, dict):
            word = html.escape(str(item.get("word") or item.get("term") or item.get("related_word") or ""))
            pos = html.escape(str(item.get("pos") or item.get("part_of_speech") or ""))
            meaning = html.escape(str(item.get("meaning") or item.get("translation") or ""))
            detail = " ".join(part for part in [word, f"[{pos}]" if pos else "", meaning] if part)
            rows.append(f"<li>{detail}</li>")
        else:
            rows.append(f"<li>{html.escape(str(item))}</li>")
    return "<ul>" + "".join(rows) + "</ul>"


def html_meanings(items: list[dict[str, str]]) -> str:
    rows = []
    for item in items:
        meaning = html.escape(item["meaning"])
        image_memory = html.escape(item["image_memory"])
        rows.append(f"<li><b>{meaning}</b><br>{image_memory}</li>")
    return "<ul>" + "".join(rows) + "</ul>"


def html_collocations(items: list[dict[str, str]]) -> str:
    rows = []
    for item in items:
        phrase = html.escape(item["phrase"])
        translation = html.escape(item["translation"])
        example = html.escape(item["example"])
        rows.append(f"<li><b>{phrase}</b>：{translation}<br><i>{example}</i></li>")
    return "<ul>" + "".join(rows) + "</ul>"


def render_image_gallery(entry: dict[str, Any], media_filenames: list[str]) -> str:
    parts: list[str] = []
    credits = entry.get("image_credits")
    if not isinstance(credits, list):
        credits = []

    for index, media_filename in enumerate(media_filenames):
        safe_filename = html.escape(media_filename, quote=True)
        parts.append(
            f'<img src="{safe_filename}" alt="{html.escape(entry["word"], quote=True)} mnemonic {index + 1}" style="max-width:100%;">'
        )
        credit = ""
        if index < len(credits):
            credit = str(credits[index]).strip()
        elif index == 0:
            credit = str(entry.get("image_credit") or "").strip()
        if credit:
            parts.append(f'<div style="font-size:12px;color:#777;">{html.escape(credit)}</div>')
    return "<br>".join(parts)


def render_full_note(entry: dict[str, Any], media_filenames: list[str]) -> str:
    if entry.get("note_html"):
        return entry["note_html"]

    parts = []
    if media_filenames:
        parts.append(render_image_gallery(entry, media_filenames))
    parts.extend(
        [
            '<div><b>常见含义和图片记忆</b></div>',
            html_meanings(entry["meanings_image"]),
            '<div><b>相关词扩展</b></div>',
            html_list(entry["related_words"]),
            '<div><b>常见搭配</b></div>',
            html_collocations(entry["collocations"]),
            '<div><b>记忆方法</b></div>',
            f'<div>{html.escape(entry["memory_method"])}</div>',
        ]
    )
    return "<br>".join(parts)


def render_sentences(entry: dict[str, Any]) -> str:
    return "<br>".join(
        f'{html.escape(item["phrase"])}: {html.escape(item["example"])}'
        for item in entry["collocations"]
    )


def build_ai_fields(entry: dict[str, Any], media_filenames: list[str]) -> dict[str, str]:
    return {
        "Word": html.escape(entry["word"]),
        "Phonetic": html.escape(entry["phonetic"]),
        "MemoryImage": render_image_gallery(entry, media_filenames) if media_filenames else "",
        "MeaningsImage": html_meanings(entry["meanings_image"]),
        "RelatedWords": html_list(entry["related_words"]),
        "Collocations": html_collocations(entry["collocations"]),
        "MemoryMethod": html.escape(entry["memory_method"]),
    }


def build_configured_fields(entry: dict[str, Any], media_filenames: list[str], config: dict[str, Any]) -> dict[str, str]:
    field_map = config.get("field_map")
    if not field_map:
        return build_ai_fields(entry, media_filenames)

    source_urls = entry.get("image_source_urls")
    if isinstance(source_urls, list):
        rendered_source_urls = "<br>".join(html.escape(str(url)) for url in source_urls if str(url).strip())
    else:
        rendered_source_urls = html.escape(entry.get("image_source_url", ""))

    values = {
        "word": html.escape(entry["word"]),
        "phonetic": html.escape(entry["phonetic"]),
        "glossary": render_full_note(entry, media_filenames),
        "sentences": render_sentences(entry),
        "note": "<br>".join(
            [
                "<b>相关词扩展</b>",
                html_list(entry["related_words"]),
                "<b>记忆方法</b>",
                html.escape(entry["memory_method"]),
            ]
        ),
        "full_note": render_full_note(entry, media_filenames),
        "source_url": rendered_source_urls,
        "audio": "",
        "image": render_image_gallery(entry, media_filenames),
        "empty": "",
    }

    fields: dict[str, str] = {}
    for anki_field, source_key in field_map.items():
        fields[anki_field] = values.get(source_key, "")
    return fields


def anki_query_escape(value: str) -> str:
    return value.replace("\\", "\\\\").replace('"', '\\"')


def note_exists(deck: str, note_type: str, fields: dict[str, str], config: dict[str, Any]) -> bool:
    duplicate = config.get("duplicate_check", {})
    if duplicate is False:
        return False

    field_name = duplicate.get("field", "expression") if isinstance(duplicate, dict) else "expression"
    value = fields.get(field_name, "")
    if not value:
        return False

    query = (
        f'deck:"{anki_query_escape(deck)}" '
        f'note:"{anki_query_escape(note_type)}" '
        f'{field_name}:"{anki_query_escape(value)}"'
    )
    return bool(invoke_anki("findNotes", {"query": query}))


def existing_note_ids_for_word(deck: str, note_type: str, word: str, config: dict[str, Any]) -> list[int]:
    duplicate = config.get("duplicate_check", {})
    if duplicate is False:
        return []

    field_name = duplicate.get("field", "expression") if isinstance(duplicate, dict) else "expression"
    query = (
        f'deck:"{anki_query_escape(deck)}" '
        f'note:"{anki_query_escape(note_type)}" '
        f'{field_name}:"{anki_query_escape(word)}"'
    )
    note_ids = invoke_anki("findNotes", {"query": query})
    return [int(note_id) for note_id in note_ids]


def add_note(deck: str, note_type: str, entry: dict[str, Any], tags: list[str], media_filenames: list[str], config: dict[str, Any]) -> int:
    fields = build_configured_fields(entry, media_filenames, config)
    if note_exists(deck, note_type, fields, config):
        raise RuntimeError(f"duplicate note found for {entry['word']!r}; skipped")

    result = invoke_anki(
        "addNote",
        {
            "note": {
                "deckName": deck,
                "modelName": note_type,
                "fields": fields,
                "options": {
                    "allowDuplicate": False,
                    "duplicateScope": "deck",
                },
                "tags": tags,
            }
        },
    )
    return int(result)


def write_preview(path: str, entries: list[dict[str, Any]]) -> None:
    with open(path, "w", encoding="utf-8") as handle:
        json.dump(entries, handle, ensure_ascii=False, indent=2)


def success_summary_line(word: str, ai_status: str, anki_status: str, detail: str = "", note_id: Any = "") -> str:
    parts = [word, f"AI={ai_status}", f"ANKI={anki_status}"]
    if detail:
        parts.append(detail)
    if note_id != "":
        parts.append(f"note_id={note_id}")
    return "\t".join(parts)


def failed_summary_line(word: str, ai_status: str, anki_status: str, reason: str) -> str:
    return "\t".join([word, f"AI={ai_status}", f"ANKI={anki_status}", reason])


def print_word_log_block(stage: str, index: int, total: int, word: str, detail: str = "") -> None:
    line = "=" * 88
    print(line, file=sys.stderr, flush=True)
    message = f"[{index}/{total}] {stage}: {word}"
    if detail:
        message += f" | {detail}"
    print(message, file=sys.stderr, flush=True)


def print_run_log_block(title: str, detail: str = "") -> None:
    line = "=" * 88
    print(line, file=sys.stderr, flush=True)
    print(title, file=sys.stderr, flush=True)
    if detail:
        print(detail, file=sys.stderr, flush=True)


def process_word_job(
    index: int,
    total: int,
    word: str,
    entry_from_input: dict[str, Any] | None,
    args: argparse.Namespace,
    llm: dict[str, str],
    image_source: str,
) -> dict[str, Any]:
    started_at = time.monotonic()
    result: dict[str, Any] = {
        "index": index,
        "word": word,
        "entry": None,
        "started_at": started_at,
        "status": "ok",
        "ai_status": "pending",
        "image_status": "pending",
        "anki_status": "pending",
        "reason": "",
        "note_action": "",
        "media_filenames": [],
        "elapsed": 0.0,
    }
    try:
        if entry_from_input is None:
            entry = generate_word_entry(
                word,
                llm["model"],
                llm["api_url"],
                llm["api_key"],
                llm["prompt_template"],
            )
            result["ai_status"] = "success"
        else:
            entry = normalize_entry(entry_from_input, word)
            result["ai_status"] = "preloaded"

        result["entry"] = entry

        media_filenames: list[str] = []
        image_bytes = None
        image_settings = getattr(args, "image_settings", {"provider": "pixabay", "api_key": "", "per_word_limit": 3})
        if image_source == "internet" and not get_image_paths(entry) and not args.no_images:
            try:
                image_paths, _ = download_internet_images(
                    entry,
                    args.image_dir,
                    image_settings["provider"],
                    image_settings["api_key"],
                    image_settings["per_word_limit"],
                    image_settings["per_page"],
                )
                media_filenames = [os.path.basename(path) for path in image_paths]
                result["image_status"] = "ok"
            except Exception as exc:
                result["image_status"] = "skipped"
                result["reason"] = str(exc)
        elif get_image_paths(entry):
            media_filenames = [os.path.basename(path) for path in get_image_paths(entry)]
            result["image_status"] = "preloaded"
        elif not args.no_images:
            try:
                image_bytes = generate_memory_image(
                    entry,
                    args.image_model,
                    args.image_size,
                    args.image_quality,
                )
                image_path = save_image_file(args.image_dir, entry["word"], image_bytes)
                media_filenames = [os.path.basename(image_path)]
                entry["image_files"] = [image_path]
                entry["image_file"] = image_path
                result["image_status"] = "ok"
            except Exception as exc:
                result["image_status"] = "skipped"
                if not result["reason"]:
                    result["reason"] = str(exc)
        else:
            result["image_status"] = "disabled"

        result["media_filenames"] = media_filenames
        result["image_bytes"] = image_bytes
        return result
    except NonEnglishWordError:
        result["status"] = "failed"
        result["ai_status"] = "failed"
        result["anki_status"] = "not_created"
        result["reason"] = "非英文单词"
        return result
    except Exception as exc:
        result["status"] = "failed"
        result["ai_status"] = "failed"
        result["anki_status"] = "not_created"
        result["reason"] = str(exc)
        return result
    finally:
        result["elapsed"] = time.monotonic() - started_at


def run_threaded_mode(
    args: argparse.Namespace,
    config: dict[str, Any],
    outputs: dict[str, Any],
    input_base_dir: Path,
    input_stem: str | None,
    logger: StepLogger,
    deck: str,
    note_type: str,
    image_source: str,
    llm: dict[str, str],
    words: list[str],
    entries_to_add: list[dict[str, Any]],
) -> int:
    entries: list[dict[str, Any]] = []
    added = 0
    import_success: list[str] = []
    import_failed: list[str] = []
    import_success_rows: list[dict[str, Any]] = []
    import_failed_rows: list[dict[str, Any]] = []
    written_success_count = 0
    written_failed_count = 0
    success_txt = resolve_output_path(
        outputs.get("success_txt"),
        "import_success.txt",
        input_base_dir,
        input_stem,
        f"{input_stem}_success.txt" if input_stem else "import_success.txt",
    )
    success_xlsx = resolve_output_path(
        outputs.get("success_xlsx"),
        "import_success.xlsx",
        input_base_dir,
        input_stem,
        f"{input_stem}_success.xlsx" if input_stem else "import_success.xlsx",
    )
    failed_txt = resolve_output_path(
        outputs.get("failed_txt"),
        "import_failed.txt",
        input_base_dir,
        input_stem,
        f"{input_stem}_failed.txt" if input_stem else "import_failed.txt",
    )
    failed_xlsx = resolve_output_path(
        outputs.get("failed_xlsx"),
        "import_failed.xlsx",
        input_base_dir,
        input_stem,
        f"{input_stem}_failed.xlsx" if input_stem else "import_failed.xlsx",
    )
    preview_json_path = resolve_relative_path(args.preview_json, input_base_dir) if args.preview_json else None
    run_started_at = time.monotonic()
    if hasattr(logger, "console"):
        logger.console = False

    def flush_outputs() -> None:
        nonlocal written_success_count, written_failed_count
        logger.emit("write_import_success_txt", "running", "", str(success_txt))
        write_text(success_txt, import_success[written_success_count:])
        logger.emit("write_import_success_xlsx", "running", "", str(success_xlsx))
        write_xlsx(
            success_xlsx,
            "ImportSuccess",
            ["timestamp", "word", "ai_status", "anki_status", "action", "note_id", "deck", "note_type"],
            import_success_rows[written_success_count:],
        )
        logger.emit("write_import_failed_txt", "running", "", str(failed_txt))
        write_text(failed_txt, import_failed[written_failed_count:])
        logger.emit("write_import_failed_xlsx", "running", "", str(failed_xlsx))
        write_xlsx(
            failed_xlsx,
            "ImportFailed",
            ["timestamp", "word", "ai_status", "anki_status", "stage", "error_type", "error_message", "deck", "note_type"],
            import_failed_rows[written_failed_count:],
        )
        written_success_count = len(import_success)
        written_failed_count = len(import_failed)

    def log_final_word(result: dict[str, Any], anki_status: str, action: str = "", reason: str = "") -> None:
        elapsed = time.monotonic() - result["started_at"]
        summary = f"AI={result['ai_status']} IMAGE={result['image_status']} ANKI={anki_status}"
        if action:
            summary += f" ACTION={action}"
        if reason:
            summary += f" REASON={reason}"
        summary += f" ELAPSED={elapsed:.1f}s"
        print_word_log_block("END", result["index"], total, result["word"], summary)
        logger.emit("word_end", "ok", result["word"], summary)

    def log_ready_word(result: dict[str, Any], completed: int, pending_total: int) -> None:
        summary = (
            f"progress={completed}/{pending_total} "
            f"AI={result['ai_status']} IMAGE={result['image_status']} "
            f"WORKER={result['status']} ELAPSED={result['elapsed']:.1f}s"
        )
        if result.get("reason"):
            summary += f" NOTE={result['reason']}"
        print_word_log_block("READY", result["index"], total, result["word"], summary)
        logger.emit("word_ready", result["status"], result["word"], summary)

    print_run_log_block(
        "RUN START",
        f"total={len(words)} deck={deck} note_type={note_type} source={'entries_json' if args.entries_json else 'words'} threads={args.threads}",
    )

    pending_jobs: list[dict[str, Any]] = []
    total = len(words)
    for index, word in enumerate(words, start=1):
        if not args.entries_json:
            existing_note_ids = existing_note_ids_for_word(deck, note_type, word, config)
            if existing_note_ids:
                reason = "Anki 已存在该单词"
                logger.emit("precheck_duplicate", "skipped", word, reason)
                print_word_log_block("SKIP", index, total, word, reason)
                import_failed.append(failed_summary_line(word, "not_created", "duplicate", reason))
                import_failed_rows.append(
                    {
                        "timestamp": dt.datetime.now().isoformat(timespec="seconds"),
                        "word": word,
                        "ai_status": "not_created",
                        "anki_status": "duplicate",
                        "stage": "precheck_duplicate",
                        "error_type": "DuplicateNoteError",
                        "error_message": reason,
                        "deck": deck,
                        "note_type": note_type,
                    }
                )
                continue

        entry_from_input = entries_to_add[index - 1] if args.entries_json else None
        pending_jobs.append({"index": index, "word": word, "entry": entry_from_input})

    if not pending_jobs:
        flush_outputs()
        if preview_json_path:
            logger.emit("write_preview_json", "running", "", str(preview_json_path))
            write_preview(str(preview_json_path), entries)
        print_run_log_block(
            "RUN END",
            f"entries={len(entries)} imported={added} failed={len(import_failed)} elapsed={time.monotonic() - run_started_at:.1f}s",
        )
        logger.emit("run_done", "ok", "", f"entries={len(entries)} imported={added} failed={len(import_failed)}")
        return 0

    print_run_log_block(
        "THREAD WORKERS START",
        f"pending={len(pending_jobs)} threads={args.threads} ai_and_image=parallel anki_write=serial",
    )
    logger.emit(
        "thread_workers_start",
        "running",
        "",
        f"pending={len(pending_jobs)} threads={args.threads} anki_write=serial",
    )

    with ThreadPoolExecutor(max_workers=args.threads) as executor:
        futures = [
            executor.submit(
                process_word_job,
                job["index"],
                total,
                job["word"],
                job["entry"],
                args,
                llm,
                image_source,
            )
            for job in pending_jobs
        ]

        completed_jobs = 0
        for future in as_completed(futures):
            result = future.result()
            completed_jobs += 1
            log_ready_word(result, completed_jobs, len(pending_jobs))
            entry = result["entry"]
            if result["status"] != "ok" or entry is None:
                import_failed.append(
                    failed_summary_line(
                        result["word"],
                        result.get("ai_status", "failed"),
                        result.get("anki_status", "not_created"),
                        result.get("reason", "unknown error"),
                    )
                )
                import_failed_rows.append(
                    {
                        "timestamp": dt.datetime.now().isoformat(timespec="seconds"),
                        "word": result["word"],
                        "ai_status": result.get("ai_status", "failed"),
                        "anki_status": result.get("anki_status", "not_created"),
                        "stage": "generate_entry",
                        "error_type": "NonEnglishWordError" if result.get("reason") == "非英文单词" else "WorkerError",
                        "error_message": result.get("reason", "unknown error"),
                        "deck": deck,
                        "note_type": note_type,
                    }
                )
                flush_outputs()
                log_final_word(result, result.get("anki_status", "not_created"), reason=result.get("reason", "unknown error"))
                continue

            entries.append(entry)
            media_filenames = result.get("media_filenames") or []
            image_bytes = result.get("image_bytes")

            if args.dry_run:
                logger.emit("dry_run_skip_import", "ok", entry["word"])
                result["anki_status"] = "dry_run"
                import_success.append(success_summary_line(entry["word"], result["ai_status"], "dry_run", "dry_run", ""))
                import_success_rows.append(
                    {
                        "timestamp": dt.datetime.now().isoformat(timespec="seconds"),
                        "word": entry["word"],
                        "ai_status": result["ai_status"],
                        "anki_status": "dry_run",
                        "action": "dry_run",
                        "note_id": "",
                        "deck": deck,
                        "note_type": note_type,
                    }
                )
                log_final_word(result, "dry_run", reason=result.get("reason", ""))
                flush_outputs()
                continue

            image_paths = get_image_paths(entry)
            if image_paths:
                try:
                    logger.emit("store_media", "running", entry["word"], ", ".join(image_paths))
                    media_filenames = store_anki_media_from_paths(image_paths)
                    logger.emit("store_media", "ok", entry["word"], ", ".join(media_filenames))
                except RuntimeError as exc:
                    reason = str(exc)
                    result["anki_status"] = "failed"
                    import_failed.append(failed_summary_line(entry["word"], "success", "failed", reason))
                    import_failed_rows.append(
                        {
                            "timestamp": dt.datetime.now().isoformat(timespec="seconds"),
                            "word": entry["word"],
                            "ai_status": "success",
                            "anki_status": "failed",
                            "stage": "store_media",
                            "error_type": type(exc).__name__,
                            "error_message": reason,
                            "deck": deck,
                            "note_type": note_type,
                        }
                    )
                    logger.emit("store_media", "failed", entry["word"], reason)
                    flush_outputs()
                    log_final_word(result, "failed", reason=reason)
                    continue
            elif media_filenames and image_bytes:
                try:
                    logger.emit("store_media", "running", entry["word"], ", ".join(media_filenames))
                    store_anki_media(media_filenames[0], image_bytes)
                    logger.emit("store_media", "ok", entry["word"], ", ".join(media_filenames))
                except RuntimeError as exc:
                    reason = str(exc)
                    result["anki_status"] = "failed"
                    import_failed.append(failed_summary_line(entry["word"], "success", "failed", reason))
                    import_failed_rows.append(
                        {
                            "timestamp": dt.datetime.now().isoformat(timespec="seconds"),
                            "word": entry["word"],
                            "ai_status": "success",
                            "anki_status": "failed",
                            "stage": "store_media",
                            "error_type": type(exc).__name__,
                            "error_message": reason,
                            "deck": deck,
                            "note_type": note_type,
                        }
                    )
                    logger.emit("store_media", "failed", entry["word"], reason)
                    flush_outputs()
                    log_final_word(result, "failed", reason=reason)
                    continue

            logger.emit("add_or_update_note", "running", entry["word"])
            print_word_log_block("ANKI", result["index"], total, entry["word"], "writing note serially")
            try:
                note_id = add_note(deck, note_type, entry, args.tag, media_filenames, config)
                added += 1
                result["anki_status"] = "success"
                result["note_action"] = f"added note_id={note_id}"
                import_success.append(success_summary_line(entry["word"], result["ai_status"], "success", "added", note_id))
                import_success_rows.append(
                    {
                        "timestamp": dt.datetime.now().isoformat(timespec="seconds"),
                        "word": entry["word"],
                        "ai_status": result["ai_status"],
                        "anki_status": "success",
                        "action": "added",
                        "note_id": note_id,
                        "deck": deck,
                        "note_type": note_type,
                    }
                )
                logger.emit("add_or_update_note", "ok", entry["word"])
            except RuntimeError as exc:
                reason = str(exc)
                result["anki_status"] = "failed"
                import_failed.append(failed_summary_line(entry["word"], result["ai_status"], "failed", reason))
                import_failed_rows.append(
                    {
                        "timestamp": dt.datetime.now().isoformat(timespec="seconds"),
                        "word": entry["word"],
                        "ai_status": result["ai_status"],
                        "anki_status": "failed",
                        "stage": "add_or_update_note",
                        "error_type": type(exc).__name__,
                        "error_message": reason,
                        "deck": deck,
                        "note_type": note_type,
                    }
                )
                logger.emit("add_or_update_note", "failed", entry["word"], reason)

            log_final_word(result, result["anki_status"], action=result["note_action"], reason=result.get("reason", ""))
            flush_outputs()
            if args.sleep:
                time.sleep(args.sleep)

    flush_outputs()
    if preview_json_path:
        logger.emit("write_preview_json", "running", "", str(preview_json_path))
        write_preview(str(preview_json_path), entries)

    print_run_log_block(
        "RUN END",
        f"entries={len(entries)} imported={added} failed={len(import_failed)} elapsed={time.monotonic() - run_started_at:.1f}s",
    )
    logger.emit("run_done", "ok", "", f"entries={len(entries)} imported={added} failed={len(import_failed)}")
    return 0


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate AI English word cards and add them to Anki through AnkiConnect."
    )
    parser.add_argument("words", nargs="*", help="Words to add. Also accepts comma-separated input.")
    parser.add_argument("-f", "--file", help="Text file containing words separated by newlines, spaces, or commas.")
    parser.add_argument("--entries-json", help="Use pre-generated entries JSON instead of calling OpenAI.")
    parser.add_argument("--config", default=DEFAULT_CONFIG, help=f"JSON config file. Default: {DEFAULT_CONFIG}")
    parser.add_argument("--deck", default=DEFAULT_DECK, help=f"Anki deck name. Default: {DEFAULT_DECK}")
    parser.add_argument("--note-type", default=DEFAULT_NOTE_TYPE, help=f"Anki note type. Default: {DEFAULT_NOTE_TYPE}")
    parser.add_argument(
        "--provider",
        choices=["openai", "doubao"],
        help=f"LLM provider for word generation. Default: {DEFAULT_LLM_PROVIDER}",
    )
    parser.add_argument("--model", help="LLM model override.")
    parser.add_argument("--api-url", help="LLM responses API URL override.")
    parser.add_argument("--image-model", default=DEFAULT_IMAGE_MODEL, help=f"OpenAI image model. Default: {DEFAULT_IMAGE_MODEL}")
    parser.add_argument("--image-size", default="1024x1024", help="Image size. Default: 1024x1024")
    parser.add_argument("--image-quality", default="low", help="Image quality: low, medium, or high. Default: low")
    parser.add_argument("--image-dir", default="generated_images", help="Where generated PNG files are saved locally.")
    parser.add_argument("--no-images", action="store_true", help="Skip real image generation.")
    parser.add_argument("--threads", type=int, help="Enable threaded mode with this many workers for AI and image steps.")
    parser.add_argument("--tag", action="append", default=["ai-word"], help="Anki tag. Can be repeated.")
    parser.add_argument("--dry-run", action="store_true", help="Generate entries but do not write to Anki.")
    parser.add_argument("--preview-json", help="Write generated entries to this JSON file.")
    parser.add_argument("--sleep", type=float, default=0.2, help="Seconds to wait between words.")
    parser.add_argument("--update-existing", action="store_true", help="Update existing matching notes instead of adding new ones.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    config = load_config(args.config)
    args.image_settings = resolve_image_settings(config)
    outputs = config.get("outputs", {})
    input_base_dir = Path.cwd()
    input_stem: str | None = None
    if args.file:
        input_path = Path(args.file).resolve()
        input_base_dir = input_path.parent
        input_stem = input_path.stem
    logger = StepLogger(
        resolve_output_path(
            outputs.get("log_jsonl"),
            "import_log.jsonl",
            input_base_dir,
            input_stem,
            f"{input_stem}_log.jsonl" if input_stem else "import_log.jsonl",
        )
    )
    logger.emit("config_loaded", "ok", "", args.config)
    deck = args.deck if args.deck != DEFAULT_DECK else config.get("deck", args.deck)
    note_type = args.note_type if args.note_type != DEFAULT_NOTE_TYPE else config.get("note_type", args.note_type)
    image_source = args.image_settings["source"]
    llm = resolve_llm_settings(args, config)
    logger.emit("llm_config", "ok", "", f"{llm['provider']} {llm['model']}")

    if args.entries_json:
        entries_to_add = load_entries(args.entries_json)
        words = [entry["word"] for entry in entries_to_add]
    else:
        entries_to_add = []
        words = read_words(args)

    if not words:
        print("No words provided. Pass words as arguments, with --file, or through stdin.", file=sys.stderr)
        return 2

    if not args.dry_run:
        logger.emit("ensure_deck", "running", "", deck)
        ensure_deck(deck)
        logger.emit("ensure_deck", "ok", "", deck)
        if not config.get("use_existing_note_type", False):
            logger.emit("ensure_note_type", "running", "", note_type)
            ensure_note_type(note_type)
            logger.emit("ensure_note_type", "ok", "", note_type)

    if args.threads and args.threads > 0:
        return run_threaded_mode(
            args,
            config,
            outputs,
            input_base_dir,
            input_stem,
            logger,
            deck,
            note_type,
            image_source,
            llm,
            words,
            entries_to_add,
        )

    print_run_log_block(
        "RUN START",
        f"total={len(words)} deck={deck} note_type={note_type} source={'entries_json' if args.entries_json else 'words'}",
    )

    entries: list[dict[str, Any]] = []
    added = 0
    import_success: list[str] = []
    import_failed: list[str] = []
    import_success_rows: list[dict[str, Any]] = []
    import_failed_rows: list[dict[str, Any]] = []
    written_success_count = 0
    written_failed_count = 0
    success_txt = resolve_output_path(
        outputs.get("success_txt"),
        "import_success.txt",
        input_base_dir,
        input_stem,
        f"{input_stem}_success.txt" if input_stem else "import_success.txt",
    )
    success_xlsx = resolve_output_path(
        outputs.get("success_xlsx"),
        "import_success.xlsx",
        input_base_dir,
        input_stem,
        f"{input_stem}_success.xlsx" if input_stem else "import_success.xlsx",
    )
    failed_txt = resolve_output_path(
        outputs.get("failed_txt"),
        "import_failed.txt",
        input_base_dir,
        input_stem,
        f"{input_stem}_failed.txt" if input_stem else "import_failed.txt",
    )
    failed_xlsx = resolve_output_path(
        outputs.get("failed_xlsx"),
        "import_failed.xlsx",
        input_base_dir,
        input_stem,
        f"{input_stem}_failed.xlsx" if input_stem else "import_failed.xlsx",
    )
    preview_json_path = resolve_relative_path(args.preview_json, input_base_dir) if args.preview_json else None
    run_started_at = time.monotonic()

    for index, word in enumerate(words, start=1):
        word_started_at = time.monotonic()
        word_ai_status = "pending"
        word_image_status = "pending"
        word_anki_status = "pending"
        word_note_action = ""
        word_reason = ""
        try:
            print_word_log_block("START", index, len(words), word)
            logger.emit("word_start", "running", word, f"{index}/{len(words)}")
            if not args.entries_json:
                existing_note_ids = existing_note_ids_for_word(deck, note_type, word, config)
                if existing_note_ids:
                    reason = "Anki 已存在该单词"
                    logger.emit("precheck_duplicate", "skipped", word, reason)
                    print(f"[{index}/{len(words)}] skipped {word}: {reason}", file=sys.stderr)
                    word_ai_status = "not_created"
                    word_anki_status = "duplicate"
                    word_reason = reason
                    import_failed.append(failed_summary_line(word, "not_created", "duplicate", reason))
                    import_failed_rows.append(
                        {
                            "timestamp": dt.datetime.now().isoformat(timespec="seconds"),
                            "word": word,
                            "ai_status": "not_created",
                            "anki_status": "duplicate",
                            "stage": "precheck_duplicate",
                            "error_type": "DuplicateNoteError",
                            "error_message": reason,
                            "deck": deck,
                            "note_type": note_type,
                        }
                    )
                    continue
            if args.entries_json:
                entry = entries_to_add[index - 1]
                word_ai_status = "preloaded"
                if get_image_paths(entry):
                    word_image_status = "preloaded"
            else:
                logger.emit("generate_entry", "running", word)
                print(f"[{index}/{len(words)}] generating {word}...", file=sys.stderr)
                try:
                    entry = generate_word_entry(
                        word,
                        llm["model"],
                        llm["api_url"],
                        llm["api_key"],
                        llm["prompt_template"],
                    )
                    word_ai_status = "success"
                except NonEnglishWordError:
                    logger.emit("generate_entry", "skipped", word, "非英文单词")
                    print(f"[{index}/{len(words)}] 非英文单词: {word}", file=sys.stderr)
                    word_ai_status = "failed"
                    word_anki_status = "not_created"
                    word_reason = "非英文单词"
                    import_failed.append(failed_summary_line(word, "failed", "not_created", "非英文单词"))
                    import_failed_rows.append(
                        {
                            "timestamp": dt.datetime.now().isoformat(timespec="seconds"),
                            "word": word,
                            "ai_status": "failed",
                            "anki_status": "not_created",
                            "stage": "generate_entry",
                            "error_type": "NonEnglishWordError",
                            "error_message": "非英文单词",
                            "deck": deck,
                            "note_type": note_type,
                        }
                    )
                    continue
                except Exception as exc:
                    logger.emit("generate_entry", "failed", word, str(exc))
                    print(f"[{index}/{len(words)}] failed to generate {word}: {exc}", file=sys.stderr)
                    word_ai_status = "failed"
                    word_anki_status = "not_created"
                    word_reason = str(exc)
                    import_failed.append(failed_summary_line(word, "failed", "not_created", str(exc)))
                    import_failed_rows.append(
                        {
                            "timestamp": dt.datetime.now().isoformat(timespec="seconds"),
                            "word": word,
                            "ai_status": "failed",
                            "anki_status": "not_created",
                            "stage": "generate_entry",
                            "error_type": type(exc).__name__,
                            "error_message": str(exc),
                            "deck": deck,
                            "note_type": note_type,
                        }
                    )
                    continue
                logger.emit("generate_entry", "ok", word)
            entries.append(entry)

            media_filenames: list[str] = []
            image_bytes = None
            if image_source == "internet" and not get_image_paths(entry) and not args.no_images:
                logger.emit("fetch_image", "running", entry["word"])
                print(f"[{index}/{len(words)}] fetching internet image for {entry['word']}...", file=sys.stderr)
                try:
                    image_paths, _ = download_internet_images(
                        entry,
                        args.image_dir,
                        args.image_settings["provider"],
                        args.image_settings["api_key"],
                        args.image_settings["per_word_limit"],
                        args.image_settings["per_page"],
                    )
                    media_filenames = [os.path.basename(path) for path in image_paths]
                    logger.emit("fetch_image", "ok", entry["word"], ", ".join(image_paths))
                    word_image_status = "ok"
                except Exception as exc:
                    logger.emit("fetch_image", "skipped", entry["word"], str(exc))
                    print(f"[{index}/{len(words)}] skip image for {entry['word']}: {exc}", file=sys.stderr)
                    word_image_status = "skipped"
                    word_reason = str(exc) if not word_reason else word_reason
            elif args.entries_json and get_image_paths(entry):
                media_filenames = [os.path.basename(path) for path in get_image_paths(entry)]
            elif not args.no_images:
                logger.emit("generate_image", "running", entry["word"])
                print(f"[{index}/{len(words)}] generating image for {entry['word']}...", file=sys.stderr)
                try:
                    image_bytes = generate_memory_image(
                        entry,
                        args.image_model,
                        args.image_size,
                        args.image_quality,
                    )
                    image_path = save_image_file(args.image_dir, entry["word"], image_bytes)
                    media_filenames = [os.path.basename(image_path)]
                    entry["image_files"] = [image_path]
                    entry["image_file"] = image_path
                    logger.emit("generate_image", "ok", entry["word"], image_path)
                    word_image_status = "ok"
                except Exception as exc:
                    logger.emit("generate_image", "skipped", entry["word"], str(exc))
                    print(f"[{index}/{len(words)}] skip image for {entry['word']}: {exc}", file=sys.stderr)
                    word_image_status = "skipped"
                    word_reason = str(exc) if not word_reason else word_reason
            else:
                word_image_status = "disabled"

            if args.dry_run:
                logger.emit("dry_run_skip_import", "ok", entry["word"])
                word_anki_status = "dry_run"
                import_success.append(success_summary_line(entry["word"], word_ai_status, "dry_run", "dry_run", ""))
                import_success_rows.append(
                    {
                        "timestamp": dt.datetime.now().isoformat(timespec="seconds"),
                        "word": entry["word"],
                        "ai_status": word_ai_status,
                        "anki_status": "dry_run",
                        "action": "dry_run",
                        "note_id": "",
                        "deck": deck,
                        "note_type": note_type,
                    }
                )
                continue

            image_paths = get_image_paths(entry)
            if image_paths:
                logger.emit("store_media", "running", entry["word"], ", ".join(image_paths))
                media_filenames = store_anki_media_from_paths(image_paths)
                logger.emit("store_media", "ok", entry["word"], ", ".join(media_filenames))
            elif media_filenames and image_bytes:
                logger.emit("store_media", "running", entry["word"], ", ".join(media_filenames))
                store_anki_media(media_filenames[0], image_bytes)
                logger.emit("store_media", "ok", entry["word"], ", ".join(media_filenames))

            logger.emit("add_or_update_note", "running", entry["word"])
            print(f"[{index}/{len(words)}] adding {entry['word']} to Anki...", file=sys.stderr)
            try:
                fields = build_configured_fields(entry, media_filenames, config)
                duplicate = config.get("duplicate_check", {})
                duplicate_field = duplicate.get("field", "expression") if isinstance(duplicate, dict) else "expression"
                query = (
                    f'deck:"{anki_query_escape(deck)}" '
                    f'note:"{anki_query_escape(note_type)}" '
                    f'{duplicate_field}:"{anki_query_escape(fields.get(duplicate_field, ""))}"'
                )
                existing_notes = invoke_anki("findNotes", {"query": query})
                if existing_notes:
                    if args.update_existing:
                        for note_id in existing_notes:
                            invoke_anki("updateNoteFields", {"note": {"id": note_id, "fields": fields}})
                            print(f"updated {entry['word']} note_id={note_id}", file=sys.stderr)
                            word_anki_status = "updated"
                            word_note_action = "updated"
                            import_success_rows.append(
                                {
                                    "timestamp": dt.datetime.now().isoformat(timespec="seconds"),
                                    "word": entry["word"],
                                    "ai_status": "success",
                                    "anki_status": "success",
                                    "action": "updated",
                                    "note_id": note_id,
                                    "deck": deck,
                                    "note_type": note_type,
                                }
                            )
                        added += len(existing_notes)
                        import_success.append(success_summary_line(entry["word"], "success", "success", "updated", existing_notes[-1]))
                    else:
                        reason = "Anki 已存在该单词"
                        print(f"skipped {entry['word']}: {reason}", file=sys.stderr)
                        word_anki_status = "duplicate"
                        word_reason = reason
                        import_failed.append(failed_summary_line(entry["word"], "success", "duplicate", reason))
                        import_failed_rows.append(
                            {
                                "timestamp": dt.datetime.now().isoformat(timespec="seconds"),
                                "word": entry["word"],
                                "ai_status": "success",
                                "anki_status": "duplicate",
                                "stage": "duplicate_check",
                                "error_type": "DuplicateNoteError",
                                "error_message": reason,
                                "deck": deck,
                                "note_type": note_type,
                            }
                        )
                        logger.emit("add_or_update_note", "skipped", entry["word"], reason)
                        continue
                else:
                    note_id = add_note(deck, note_type, entry, args.tag, media_filenames, config)
                    added += 1
                    print(f"added {entry['word']} note_id={note_id}", file=sys.stderr)
                    word_anki_status = "success"
                    word_note_action = "added"
                    import_success.append(success_summary_line(entry["word"], "success", "success", "added", note_id))
                    import_success_rows.append(
                        {
                            "timestamp": dt.datetime.now().isoformat(timespec="seconds"),
                            "word": entry["word"],
                            "ai_status": "success",
                            "anki_status": "success",
                            "action": "added",
                            "note_id": note_id,
                            "deck": deck,
                            "note_type": note_type,
                        }
                    )
                logger.emit("add_or_update_note", "ok", entry["word"])
            except RuntimeError as exc:
                print(f"skipped {entry['word']}: {exc}", file=sys.stderr)
                word_anki_status = "failed"
                word_reason = str(exc)
                import_failed.append(failed_summary_line(entry["word"], "success", "failed", str(exc)))
                import_failed_rows.append(
                    {
                        "timestamp": dt.datetime.now().isoformat(timespec="seconds"),
                        "word": entry["word"],
                        "ai_status": "success",
                        "anki_status": "failed",
                        "stage": "add_or_update_note",
                        "error_type": type(exc).__name__,
                        "error_message": str(exc),
                        "deck": deck,
                        "note_type": note_type,
                    }
                )
                logger.emit("add_or_update_note", "failed", entry["word"], str(exc))
            if args.sleep:
                time.sleep(args.sleep)
        finally:
            elapsed = time.monotonic() - word_started_at
            summary = f"AI={word_ai_status} IMAGE={word_image_status} ANKI={word_anki_status}"
            if word_note_action:
                summary += f" ACTION={word_note_action}"
            if word_reason:
                summary += f" REASON={word_reason}"
            summary += f" ELAPSED={elapsed:.1f}s"
            print_word_log_block("END", index, len(words), word, summary)
            logger.emit("word_end", "ok", word, summary)

        logger.emit("write_import_success_txt", "running", "", str(success_txt))
        write_text(success_txt, import_success[written_success_count:])
        logger.emit("write_import_success_xlsx", "running", "", str(success_xlsx))
        write_xlsx(
            success_xlsx,
            "ImportSuccess",
            ["timestamp", "word", "ai_status", "anki_status", "action", "note_id", "deck", "note_type"],
            import_success_rows[written_success_count:],
        )
        logger.emit("write_import_failed_txt", "running", "", str(failed_txt))
        write_text(failed_txt, import_failed[written_failed_count:])
        logger.emit("write_import_failed_xlsx", "running", "", str(failed_xlsx))
        write_xlsx(
            failed_xlsx,
            "ImportFailed",
            ["timestamp", "word", "ai_status", "anki_status", "stage", "error_type", "error_message", "deck", "note_type"],
            import_failed_rows[written_failed_count:],
        )
        written_success_count = len(import_success)
        written_failed_count = len(import_failed)

    if preview_json_path:
        logger.emit("write_preview_json", "running", "", str(preview_json_path))
        write_preview(str(preview_json_path), entries)

    if args.dry_run:
        print(json.dumps(entries, ensure_ascii=False, indent=2))
    else:
        print(f"Done. Generated {len(entries)} entries, added {added} notes to {deck}.")
    print_run_log_block(
        "RUN END",
        f"entries={len(entries)} imported={added} failed={len(import_failed)} elapsed={time.monotonic() - run_started_at:.1f}s",
    )
    logger.emit("run_done", "ok", "", f"entries={len(entries)} imported={added} failed={len(import_failed)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
