# ankimaker

读取一批英文单词，生成中文解释卡片和记忆图片，然后通过 AnkiConnect 写入 Anki。

默认配置见 `anki_config.json`：

- 目标牌组：`日常想到的`
- 模板：`OnlineDictHelper`
- 图片来源：互联网图片，当前使用 Pixabay
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
3. 如果你要让脚本自己直连大模型生成内容，直接在 `anki_config.json` 的 `llm.api_key` 里填写对应的 API key。
4. 如果你要从 Pixabay 拉取互联网图片，在 `anki_config.json` 的 `image.api_key` 里填写 Pixabay API key。`image.per_page` 控制查询页大小，`image.per_word_limit` 控制每个单词最终最多下载几张图，默认都为 3。

你也可以在 `llm.prompt_template` 里直接改豆包提示词。

默认会连接：

- AnkiConnect: `http://127.0.0.1:8765`
- LLM provider: `doubao`
- Doubao model: `doubao-seed-2-0-lite-260428`
- Anki deck: `English::AI Words`
- Anki note type: `AI English Word`

这些都可以用环境变量或命令行参数覆盖。

如果你想切回 OpenAI：

```bash
export LLM_PROVIDER="openai"
export OPENAI_API_KEY="你的 key"
export OPENAI_MODEL="gpt-4.1-mini"
```

也可以继续把 API key 写在 `anki_config.json` 的 `llm.api_key` 里，脚本会优先读取配置文件。

如果你用 `--file words.txt`，默认输出会写成 `words_success.*`、`words_failed.*` 和 `words_log.jsonl`，都在 `words.txt` 同目录。

## 让脚本直连大模型使用

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
export LLM_PROVIDER="doubao"
export DOUBAO_MODEL="doubao-seed-2-0-lite-260428"
export ANKI_DECK="English::AI Words"
export ANKI_NOTE_TYPE="AI English Word"
```

## 说明

使用 `OnlineDictHelper` 时，脚本只使用现有模板字段，不会修改模板。图片会优先从 Pixabay 获取并下载到本地后再写入 Anki；如果没有找到图片，会提示错误。
