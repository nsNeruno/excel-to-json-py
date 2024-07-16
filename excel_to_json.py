from tkinter import filedialog as fd
import json
import re
import pyperclip
from typing import List

import pyexcel as p

def split_keywords(keywords: str) -> List[str]:
    return re.split(r'[,|、|・]', keywords)

def generate_json(objs: list):
    result = json.dumps(
        {
            "type": "TAROT",
            "published": True,
            "dateRange": {
                "start": "2024-01-01",
                "end": "2028-12-31"
            },
            "price": 800,
            "category": "",
            "cardAmount": len(objs),
            "packIllustration": "",
            "cards": objs
        },
        indent=4,
        ensure_ascii=False
    )

    pyperclip.copy(result)
    print(f'Copied {len(objs)} item(s) to clipboard')

file_path = fd.askopenfilename()

if file_path:
    records = p.get_records(file_name=file_path)

    objs = list()

    for record in records:
        # 画像番号	タイトル	英語タイトル	キーワード	説明	レアリティ	区分
        card_number = record["画像番号"]
        jp_title = record["タイトル"]
        en_title = record["英語タイトル"]
        keywords = record["キーワード"]
        desc = record["説明"]
        rarity = record["レアリティ"]
        card_type = record["区分"]

        match card_type:
            case "タロット":
                card_type = "TAROT"
            case "ルノルマン":
                card_type = "LENORMAND"
            case "オラクル":
                card_type = "ORACLE"

        obj = {
            "index": card_number,
            "name": {
                "ja": jp_title,
                "en": en_title
            },
            "description": desc,
            "rarity": rarity,
            "type": card_type,
        }

        if keywords and isinstance(keywords, str):
            obj["keywords"] = split_keywords(keywords)
        else:
            obj["keywords"] = f'{keywords}'

        objs.append(obj)

    generate_json(objs)
    