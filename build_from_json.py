#!/usr/bin/env python3
"""JSON構成ファイルからPowerPointを生成するスクリプト"""
import json
import sys
from pathlib import Path

# generate_ppt.py の関数を再利用
from generate_ppt import create_pptx

def main():
    json_file = sys.argv[1] if len(sys.argv) > 1 else "cross_marketing_notta_structure.json"
    with open(json_file, encoding="utf-8") as f:
        data = json.load(f)

    output = Path(json_file).stem.replace("_structure", "") + "_presentation.pptx"
    create_pptx(data, output)

if __name__ == "__main__":
    main()
