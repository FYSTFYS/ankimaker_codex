#!/usr/bin/env python3
import argparse
import datetime as dt
import json
import os
import re
import subprocess
import sys
import tempfile
import time
import urllib.request
from pathlib import Path
from typing import Any

from run_utils import StepLogger, write_text, write_xlsx


ROOT = Path(__file__).resolve().parent
DEFAULT_SCHEMA = ROOT / "codex_word_schema.json"
DEFAULT_APPSERVER_CONFIG = ROOT / "codex_appserver_config.json"


class GenerationError(RuntimeError):
    def __init__(
        self,
        stage: str,
        message: str,
        command: list[str] | None = None,
        prompt: str = "",
        stdout: str = "",
        stderr: str = "",
    ) -> None:
        super().__init__(message)
        self.stage = stage
        self.message = message
        self.command = command or []
        self.prompt = prompt
        self.stdout = stdout
        self.stderr = stderr


def read_words(args: argparse.Namespace) -> list[str]:
    chunks: list[str] = []
    if args.file:
        chunks.append(Path(args.file).read_text(encoding="utf-8"))
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


def load_appserver_config(path: str) -> dict[str, Any]:
    return json.loads(Path(path).read_text(encoding="utf-8"))


def build_prompt(word: str, config: dict[str, Any]) -> str:
    return config["prompt_template"].replace("{{word}}", word)


def run_codex_exec(
    word: str,
    prompt: str,
    schema: Path,
    model: str | None,
    timeout: int,
    extra_args: list[str],
    logger: StepLogger,
) -> dict[str, Any]:
    with tempfile.TemporaryDirectory(prefix="ankimaker-codex-") as tmp:
        logger.emit("prepare_temp", "ok", word, tmp)
        output = Path(tmp) / f"{word}.json"
        cmd = [
            "codex",
            "exec",
            *extra_args,
            "-C",
            str(ROOT),
            "--output-schema",
            str(schema),
            "-o",
            str(output),
        ]
        if model:
            cmd.extend(["--model", model])
        cmd.append(prompt)
        logger.emit("build_command", "ok", word, "codex exec command prepared")

        logger.emit("codex_exec_start", "running", word, f"timeout={timeout}s")
        completed = subprocess.run(
            cmd,
            cwd=ROOT,
            text=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=timeout,
        )
        logger.emit("codex_exec_done", "ok" if completed.returncode == 0 else "failed", word, f"returncode={completed.returncode}")
        if completed.returncode != 0:
            raise GenerationError(
                "codex_exec",
                completed.stderr.strip() or completed.stdout.strip() or f"returncode={completed.returncode}",
                command=cmd,
                prompt=prompt,
                stdout=completed.stdout,
                stderr=completed.stderr,
            )
        if not output.exists():
            raise GenerationError("read_output", "Codex did not write the output JSON file.", command=cmd, prompt=prompt, stdout=completed.stdout, stderr=completed.stderr)
        logger.emit("read_output", "ok", word, str(output))
        try:
            entry = json.loads(output.read_text(encoding="utf-8"))
        except json.JSONDecodeError as exc:
            raise GenerationError("parse_json", str(exc), command=cmd, prompt=prompt, stdout=completed.stdout, stderr=completed.stderr) from exc
        logger.emit("parse_json", "ok", word, "schema-constrained JSON parsed")
        return entry


def start_app_server(config: dict[str, Any], logger: StepLogger) -> None:
    server = config["server"]
    command = server["start_command"]
    listen = server["listen"]
    log_path = ROOT / server.get("log_file", "appserver.log")
    logger.emit("appserver_start", "running", "", " ".join(command))
    log_handle = log_path.open("a", encoding="utf-8")
    subprocess.Popen(
        command,
        cwd=ROOT,
        stdout=log_handle,
        stderr=log_handle,
    )
    readyz = listen.replace("ws://", "http://").replace("wss://", "https://") + "/readyz"
    for attempt in range(30):
        try:
            with urllib.request.urlopen(readyz, timeout=2):
                logger.emit("appserver_ready", "ok", "", readyz)
                return
        except Exception:
            time.sleep(1)
    logger.emit("appserver_ready", "failed", "", readyz)


def main() -> int:
    parser = argparse.ArgumentParser(description="Generate word entries with local Codex CLI login.")
    parser.add_argument("words", nargs="*", help="Words to generate.")
    parser.add_argument("-f", "--file", help="Text file containing words.")
    parser.add_argument("--schema", default=str(DEFAULT_SCHEMA), help="Output JSON schema.")
    parser.add_argument(
        "--entries-json",
        "--output-json",
        dest="entries_json",
        default="generated_entries.json",
        help="Successful entries output JSON.",
    )
    parser.add_argument("--success-txt", default="success.txt", help="Successful words log.")
    parser.add_argument("--failed-txt", default="failed.txt", help="Failed words log.")
    parser.add_argument("--failed-xlsx", help="Failed words workbook.")
    parser.add_argument("--log-jsonl", help="JSONL step log.")
    parser.add_argument("--model", help="Optional Codex model override.")
    parser.add_argument("--timeout", type=int, default=180, help="Timeout per word in seconds.")
    parser.add_argument("--appserver-config", default=str(DEFAULT_APPSERVER_CONFIG), help="JSON config for appserver input.")
    parser.add_argument("--start-app-server", action="store_true", help="Start Codex app-server with configured listen URL before generation.")
    args = parser.parse_args()

    words = read_words(args)
    if not words:
        print("No words provided.", file=sys.stderr)
        return 2

    appserver_config = load_appserver_config(args.appserver_config)

    outputs = appserver_config["outputs"]
    log_jsonl = Path(args.log_jsonl or outputs.get("log_jsonl", "run_log.jsonl"))
    logger = StepLogger(log_jsonl)
    logger.emit("config_loaded", "ok", "", args.appserver_config)

    if args.start_app_server:
        start_app_server(appserver_config, logger)

    configured_schema = appserver_config.get("turn_start", {}).get("outputSchemaPath")
    schema = Path(args.schema if args.schema != str(DEFAULT_SCHEMA) else configured_schema or args.schema)
    if not schema.is_absolute():
        schema = ROOT / schema
    fallback = appserver_config.get("codex_exec_fallback", {})
    configured_model = args.model or fallback.get("model") or appserver_config.get("turn_start", {}).get("model") or appserver_config.get("server", {}).get("model")
    timeout = args.timeout if args.timeout != 180 else int(fallback.get("timeout_seconds", args.timeout))
    extra_args = fallback.get("extra_args", ["--skip-git-repo-check", "--sandbox", "read-only"])
    entries_json = Path(args.entries_json if args.entries_json != "generated_entries.json" else outputs["entries_json"])
    success_txt = Path(args.success_txt if args.success_txt != "success.txt" else outputs["success_txt"])
    failed_txt = Path(args.failed_txt if args.failed_txt != "failed.txt" else outputs["failed_txt"])
    failed_xlsx = Path(args.failed_xlsx or outputs.get("failed_xlsx", "failed.xlsx"))
    success_xlsx = Path(outputs.get("success_xlsx", "generation_success.xlsx"))
    failed_rows: list[dict[str, Any]] = []
    success_rows: list[dict[str, Any]] = []
    entries: list[dict[str, Any]] = []
    success: list[str] = []
    failed: list[str] = []

    for index, word in enumerate(words, start=1):
        logger.emit("word_start", "running", word, f"{index}/{len(words)}")
        try:
            prompt = build_prompt(word, appserver_config)
            logger.emit("prompt_built", "ok", word, f"chars={len(prompt)}")
            entry = run_codex_exec(word, prompt, schema, configured_model, timeout, extra_args, logger)
            entries.append(entry)
            success.append(word)
            success_rows.append(
                {
                    "timestamp": dt.datetime.now().isoformat(timespec="seconds"),
                    "word": word,
                    "phonetic": entry.get("phonetic", ""),
                    "model": configured_model or "",
                    "entries_json": str(entries_json),
                }
            )
            logger.emit("word_done", "success", word)
        except Exception as exc:
            if isinstance(exc, GenerationError):
                stage = exc.stage
                command = " ".join(exc.command)
                prompt = exc.prompt
                stdout = exc.stdout
                stderr = exc.stderr
                message = exc.message
            elif isinstance(exc, subprocess.TimeoutExpired):
                stage = "codex_exec_timeout"
                command = " ".join(exc.cmd) if isinstance(exc.cmd, list) else str(exc.cmd)
                prompt = prompt if "prompt" in locals() else ""
                stdout = exc.stdout or ""
                stderr = exc.stderr or ""
                message = f"timed out after {exc.timeout} seconds"
            else:
                stage = "unknown"
                command = ""
                prompt = prompt if "prompt" in locals() else ""
                stdout = ""
                stderr = ""
                message = str(exc)
            failed.append(f"{word}\t{stage}\t{message}")
            failed_rows.append(
                {
                    "timestamp": dt.datetime.now().isoformat(timespec="seconds"),
                    "word": word,
                    "stage": stage,
                    "error_type": type(exc).__name__,
                    "error_message": message,
                    "command": command,
                    "prompt": prompt,
                    "stdout": stdout,
                    "stderr": stderr,
                }
            )
            logger.emit("word_done", "failed", word, f"{stage}: {message}")

        logger.emit("write_entries_json", "running", "", str(entries_json))
        entries_json.write_text(json.dumps(entries, ensure_ascii=False, indent=2), encoding="utf-8")
        logger.emit("write_success_txt", "running", "", str(success_txt))
        write_text(success_txt, success)
        logger.emit("write_success_xlsx", "running", "", str(success_xlsx))
        write_xlsx(
            success_xlsx,
            "GenerationSuccess",
            ["timestamp", "word", "phonetic", "model", "entries_json"],
            success_rows,
        )
        logger.emit("write_failed_txt", "running", "", str(failed_txt))
        write_text(failed_txt, failed)
        logger.emit("write_failed_xlsx", "running", "", str(failed_xlsx))
        write_xlsx(
            failed_xlsx,
            "Failed",
            ["timestamp", "word", "stage", "error_type", "error_message", "command", "prompt", "stdout", "stderr"],
            failed_rows,
        )

    logger.emit("run_done", "ok", "", f"success={len(success)} failed={len(failed)}")
    print(f"Done. success={len(success)} failed={len(failed)}")
    return 0 if not failed else 1


if __name__ == "__main__":
    raise SystemExit(main())
