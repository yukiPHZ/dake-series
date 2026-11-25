# -*- coding: utf-8 -*-
from __future__ import annotations


SUMMARY_RULES = """以下のルールを必ず守ってください。
・原文に存在する事実だけを書く
・推測を事実として書かない
・不明確な内容は「要確認」とする
・発言者を特定しない
・人名を勝手に補完しない
・日付や金額を推測しない
・重要事項には元の時刻を付ける
・同じ内容を重複して書かない
・日本語で簡潔に書く
"""


def chunk_prompt(record_date: str, text: str) -> str:
    return f"""{SUMMARY_RULES}

対象日: {record_date}

次の業務音声ログを、後から検証できる短い中間要約にしてください。
各項目には可能な限り元の時刻を残してください。

---
{text}
"""


def final_prompt(record_date: str, chunk_summaries: list[str]) -> str:
    joined = "\n\n---\n\n".join(chunk_summaries)
    return f"""{SUMMARY_RULES}

対象日: {record_date}

以下の中間要約を統合して、Markdownで次の見出しだけを使ってください。

# {record_date} 日次要約

## 主な話題
## 決まったこと
## 確認が必要なこと
## 対応候補
## 要確認・不明瞭

空の見出しにも、該当がなければ「- なし」と書いてください。

---
{joined}
"""

