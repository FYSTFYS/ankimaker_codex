# ankimaker

读取一批英文单词，生成中文解释卡片和记忆图片，然后通过 AnkiConnect 写入 Anki。

默认配置见 `anki_config.json`：

- 目标牌组：`日常想到的`
- 模板：`OnlineDictHelper`
- 图片来源：互联网图片，当前使用 Wikimedia Commons
- 字段映射：`expression`, `reading`, `glossary`, `sentence`, `note`, `url`, `audio`

生成的内容会映射到 OnlineDictHelper：

- `expression`：单词
- `reading`：音标
- `glossary`：含义、图片、相关词、搭配、记忆方法
- `sentence`：搭配例句
- `note`：相关词和记忆方法
- `url`：图片来源页
- `audio`：暂留空

解释内容包含：

- 单词
- 音标
- 常见含义和对应的图片记忆
- 相关词扩展
- 常见搭配
- 记忆方法

## 准备

1. 安装并启动 Anki。
2. 在 Anki 里安装 AnkiConnect 插件，保持 Anki 打开。
3. 如果你要让脚本自己直连 OpenAI 生成内容，设置 OpenAI API key：

```bash
export OPENAI_API_KEY="你的 key"
```

默认会连接：

- AnkiConnect: `http://127.0.0.1:8765`
- OpenAI model: `gpt-4.1-mini`
- Anki deck: `English::AI Words`
- Anki note type: `AI English Word`

这些都可以用环境变量或命令行参数覆盖。

## 两步执行

现在流程分成两个可以单独执行的步骤。

### 1. 生成单词卡片内容

配置文件：`codex_appserver_config.json`

里面包含：

- app server 启动命令：`codex app-server --listen ws://127.0.0.1:4500`
- 使用模型：`gpt-5.5`
- 输出 schema：`codex_word_schema.json`
- prompt 模板
- 生成成功/失败/日志文件路径

运行：

```bash
python3 codex_generate_words.py --start-app-server wagon
```

也可以批量：

```bash
python3 codex_generate_words.py --file examples/words.txt
```

生成阶段输出：

- `generated_entries.json`：成功生成的完整结果
- `generation_success.txt`：生成成功的单词
- `generation_success.xlsx`：生成成功明细
- `generation_failed.txt`：生成失败摘要
- `generation_failed.xlsx`：生成失败明细，分列记录 stage、error、command、prompt、stdout、stderr
- `generation_log.jsonl`：每个功能节点的运行日志

### 2. 导入 Anki

配置文件：`anki_config.json`

运行：

```bash
python3 ankimaker.py --entries-json generated_entries.json --config anki_config.json --update-existing --no-images
```

导入阶段输出：

- `import_success.txt`：导入成功的单词
- `import_success.xlsx`：导入成功明细，包含 note id
- `import_failed.txt`：导入失败摘要
- `import_failed.xlsx`：导入失败明细
- `import_log.jsonl`：每个功能节点的运行日志

## 使用 Codex 生成内容

如果你没有 OpenAI API key，可以让 Codex 在当前会话里生成解释 JSON，然后由脚本从互联网获取图片并导入 Anki。

流程：

1. 把单词列表发给 Codex。
2. Codex 会为每个词生成解释 JSON，保存到当前项目。
3. 脚本会按 `anki_config.json` 从互联网获取图片，存到 `generated_images/` 并导入 Anki 媒体库。
4. Codex 调用：

```bash
python3 ankimaker.py --entries-json generated_entries.json --config anki_config.json
```

这个模式不需要 `OPENAI_API_KEY`。

预生成 JSON 的单词条目格式：

```json
[
  {
    "word": "abandon",
    "phonetic": "UK /əˈbændən/ US /əˈbændən/",
    "meanings_image": [
      {
        "meaning": "v. 放弃，抛弃",
        "image_memory": "一个人把沉重的背包丢在路边，轻装继续往前走。"
      }
    ],
    "related_words": ["abandoned adj. 被遗弃的", "abandonment n. 放弃；遗弃"],
    "collocations": [
      {
        "phrase": "abandon a plan",
        "translation": "放弃计划",
        "example": "They abandoned the plan after the storm."
      }
    ],
    "memory_method": "a + bandon 可联想为把一堆负担丢掉，核心义是“放弃/抛弃”。"
  }
]
```

## 让脚本直连 OpenAI 使用

直接输入单词：

```bash
python3 ankimaker.py abandon precise vivid
```

从文件读取：

```bash
python3 ankimaker.py --file examples/words.txt
```

指定牌组：

```bash
python3 ankimaker.py --deck "English::Vocabulary" --file examples/words.txt
```

先预览，不写入 Anki：

```bash
python3 ankimaker.py --dry-run --preview-json preview.json abandon precise
```

## 输入格式

文件或命令行都支持用空格、换行、英文逗号、中文逗号、分号分隔：

```text
abandon
precise, vivid
inevitable；subtle
```

## 常用配置

```bash
export ANKI_CONNECT_URL="http://127.0.0.1:8765"
export OPENAI_MODEL="gpt-4.1-mini"
export ANKI_DECK="English::AI Words"
export ANKI_NOTE_TYPE="AI English Word"
```

## 说明

使用 `OnlineDictHelper` 时，脚本只使用现有模板字段，不会修改模板。图片会优先从互联网获取；如果没有找到图片，会提示错误。
