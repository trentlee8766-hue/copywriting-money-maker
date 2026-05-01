#!/usr/bin/env python3
"""
文档分析脚本
功能：
1. 提取批注（Word文档）
2. 识别图片中的修改要求（OCR）
3. 解析纯文本修改要求
4. 统一输出标准化JSON格式

输出格式（统一）：
{
  "source_type": "comments | image | text",
  "source_file": "文件名",
  "total": 数量,
  "items": [
    {
      "id": "1",
      "type": "content_modify | content_add | content_delete | format_modify",
      "location": "所在位置",
      "content": "修改要求原文",
      "author": "作者（仅comments有）"
    }
  ]
}
"""

import os
import sys
import json
import argparse
import re
from typing import Dict, List, Any

# 添加模块路径
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def extract_comments(file_path: str) -> Dict[str, Any]:
    """
    从Word文档提取批注，附带位置信息

    Args:
        file_path: docx文件路径

    Returns:
        统一格式JSON
    """
    from docx import Document
    from docx.oxml.ns import qn

    try:
        doc = Document(file_path)

        # 建立批注ID -> 内容映射
        comment_map = {}
        for c in doc.comments:
            comment_map[c.comment_id] = {
                'author': c.author,
                'text': c.text
            }

        # 遍历段落，找到批注位置
        body = doc.element.body
        paras = body.findall('.//' + qn('w:p'))

        items = []
        item_id = 1
        for para in paras:
            comment_starts = para.findall('.//' + qn('w:commentRangeStart'))
            for cs in comment_starts:
                cid = int(cs.get(qn('w:id')))
                if cid not in comment_map:
                    continue

                # 获取段落文本
                para_text = ''
                for t in para.findall('.//' + qn('w:t')):
                    para_text += t.text or ''

                info = comment_map[cid]
                items.append({
                    'id': str(item_id),
                    'type': classify_comment(info['text']),
                    'location': para_text.strip(),
                    'content': info['text'],
                    'author': info['author']
                })
                item_id += 1

        return {
            'source_type': 'comments',
            'source_file': os.path.basename(file_path),
            'total': len(items),
            'items': items
        }

    except Exception as e:
        return {
            'source_type': 'comments',
            'source_file': os.path.basename(file_path),
            'error': str(e)
        }


def extract_from_image(file_path: str) -> Dict[str, Any]:
    """
    从图片提取修改要求（OCR）

    使用 MiniMax MCP 工具读取图片
    返回结构化JSON

    Args:
        file_path: 图片路径

    Returns:
        统一格式JSON
    """
    # 调用 MiniMax 图片识别
    # 命令: MINIMAX_API_KEY=... MINIMAX_API_HOST=https://api.minimaxi.com /c/Users/Administrator/AppData/Roaming/npm/pi-minimax-mcp understand "图片路径"
    import subprocess

    api_key = os.environ.get('MINIMAX_API_KEY', '')
    api_host = os.environ.get('MINIMAX_API_HOST', 'https://api.minimaxi.com')
    mcp_path = 'C:/Users/Administrator/AppData/Roaming/npm/pi-minimax-mcp'

    if not api_key:
        return {
            'source_type': 'image',
            'source_file': os.path.basename(file_path),
            'error': 'MINIMAX_API_KEY 环境变量未设置'
        }

    # 使用 bash 调用（Git Bash 环境）
    cmd = [
        'bash',
        '-c',
        f'MINIMAX_API_KEY="{api_key}" MINIMAX_API_HOST="{api_host}" {mcp_path} understand "{file_path}"'
    ]

    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            timeout=60
        )

        # 处理编码问题
        raw_output = result.stdout
        if isinstance(raw_output, bytes):
            for encoding in ['utf-8', 'gbk', 'latin1']:
                try:
                    ocr_text = raw_output.decode(encoding)
                    break
                except UnicodeDecodeError:
                    continue
            else:
                ocr_text = raw_output.decode('utf-8', errors='ignore')
        else:
            ocr_text = str(raw_output)

        if result.returncode != 0:
            stderr_text = ''
            if result.stderr:
                for enc in ['utf-8', 'gbk', 'latin1']:
                    try:
                        stderr_text = result.stderr.decode(enc)
                        break
                    except UnicodeDecodeError:
                        continue
            return {
                'source_type': 'image',
                'source_file': os.path.basename(file_path),
                'error': f'OCR失败: {stderr_text or result.returncode}'
            }

        # 图片OCR返回的是叙述性文本，需要AI理解
        # 因此同时返回原始文本供AI分析
        result = parse_text_to_items(ocr_text, os.path.basename(file_path), 'image')
        result['raw_text'] = ocr_text
        return result

    except subprocess.TimeoutExpired:
        return {
            'source_type': 'image',
            'source_file': os.path.basename(file_path),
            'error': 'OCR超时'
        }
    except Exception as e:
        return {
            'source_type': 'image',
            'source_file': os.path.basename(file_path),
            'error': str(e)
        }


def extract_from_text(text: str, source_name: str = "text_input") -> Dict[str, Any]:
    """
    从纯文本提取修改要求

    Args:
        text: 文本内容
        source_name: 来源名称

    Returns:
        统一格式JSON
    """
    return parse_text_to_items(text, source_name, 'text')


def parse_text_to_items(text: str, source_name: str, source_type: str) -> Dict[str, Any]:
    """
    将文本解析为结构化items

    Args:
        text: 原始文本
        source_name: 来源名称
        source_type: 'image' 或 'text'

    Returns:
        统一格式JSON
    """
    items = []
    item_id = 1

    # 预处理：统一换行符
    text = text.replace('\r\n', '\n').replace('\r', '\n')

    # 先尝试按"数字+点"分割（1. 2. 3.）
    # 再按句号/分号分割
    # 先用正则提取所有 "数字. 内容" 模式的片段
    numbered_pattern = r'(\d+[.、])\s*([^。；\n]+[。；]?)'

    matches = list(re.finditer(numbered_pattern, text))
    if matches and len(matches) >= 2:
        # 有多个编号项，按编号分割
        for match in matches:
            content = match.group(2).strip()
            if content:
                items.append({
                    'id': str(item_id),
                    'type': classify_text(content),
                    'location': source_name,
                    'content': content,
                    'author': ''
                })
                item_id += 1
    else:
        # 没有清晰的编号，按句子分割
        # 按 。 ！ ？ 分割
        sentences = re.split(r'([。！？\n])', text)
        current = ''
        for i, part in enumerate(sentences):
            if i % 2 == 0:
                # 文本部分
                current = part.strip()
            else:
                # 标点部分，合并
                if current:
                    full_sentence = current + part
                    full_sentence = full_sentence.strip()
                    if full_sentence and len(full_sentence) > 3:
                        items.append({
                            'id': str(item_id),
                            'type': classify_text(full_sentence),
                            'location': source_name,
                            'content': full_sentence,
                            'author': ''
                        })
                        item_id += 1
                    current = ''

        # 如果没有分句成功（没有标点），直接按行分割
        if not items:
            for line in text.split('\n'):
                line = line.strip()
                if line and len(line) > 3:
                    items.append({
                        'id': str(item_id),
                        'type': classify_text(line),
                        'location': source_name,
                        'content': line,
                        'author': ''
                    })
                    item_id += 1

    return {
        'source_type': source_type,
        'source_file': source_name,
        'total': len(items),
        'items': items
    }


def classify_comment(text: str) -> str:
    """
    根据批注内容分类

    Args:
        text: 批注文本

    Returns:
        类型: content_modify | content_add | content_delete | format_modify
    """
    text = text.strip()

    # 格式类关键词
    format_keywords = ['格式', '三线表', '表格', '签名', '日期', '对齐', '字体', '字号', '行距', '缩进']
    for kw in format_keywords:
        if kw in text:
            return 'format_modify'

    # 删除类关键词
    delete_keywords = ['删除', '去掉', '移除', '不要', '无意义', '多余', '删']
    for kw in delete_keywords:
        if kw in text:
            return 'content_delete'

    # 补充类关键词
    add_keywords = ['添加', '补充', '增加', '加入', '充实', '丰富', '需要', '要有', '加上']
    for kw in add_keywords:
        if kw in text:
            return 'content_add'

    # 默认：内容修改
    return 'content_modify'


def classify_text(text: str) -> str:
    """
    根据文本内容分类（用于image和text来源）

    Args:
        text: 文本内容

    Returns:
        类型
    """
    return classify_comment(text)  # 共用同一个分类逻辑


def main():
    parser = argparse.ArgumentParser(
        description="文档分析脚本 - 提取批注/修改要求，统一JSON输出"
    )
    parser.add_argument("--input", "-i", type=str, help="输入文件路径（docx/image）")
    parser.add_argument("--text", "-t", type=str, help="直接输入文本内容")
    parser.add_argument("--name", "-n", type=str, default="input", help="来源名称（用于text模式）")
    parser.add_argument("--output", "-o", type=str, help="输出JSON文件路径")
    parser.add_argument("--pretty", "-p", action="store_true", help="格式化输出")

    args = parser.parse_args()

    # 确定输入模式
    if args.text:
        result = extract_from_text(args.text, args.name)
    elif args.input:
        ext = os.path.splitext(args.input)[1].lower()
        if ext == '.docx':
            result = extract_comments(args.input)
        elif ext in ['.png', '.jpg', '.jpeg', '.bmp', '.tiff', '.gif']:
            result = extract_from_image(args.input)
        else:
            result = {'error': f'不支持的文件类型: {ext}'}
    else:
        print("请指定 --input 或 --text")
        return

    # 输出
    if args.pretty:
        output = json.dumps(result, ensure_ascii=False, indent=2)
    else:
        output = json.dumps(result, ensure_ascii=False)

    if args.output:
        with open(args.output, 'w', encoding='utf-8') as f:
            f.write(output)
        print(f"已保存到: {args.output}")
    else:
        print(output)

    return result


if __name__ == "__main__":
    main()
