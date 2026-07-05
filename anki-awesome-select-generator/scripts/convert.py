#!/usr/bin/env python3
"""
Anki题目导入文件生成器
支持xlsx、toml、csv、txt格式输入，输出Anki可导入的TSV文件
"""

import argparse
import os
import sys


def parse_xlsx(file_path):
    """解析xlsx文件"""
    try:
        import openpyxl
    except ImportError:
        print("需要安装openpyxl: pip install openpyxl")
        sys.exit(1)
    
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active
    
    questions = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[0]:
            continue
        questions.append({
            'number': row[0],
            'type': row[1],
            'stem': row[2] or '',
            'options_raw': row[3] or '',
            'answer': row[4] or '',
            'notes': row[5] or '' if len(row) > 5 else ''
        })
    return questions


def parse_toml(file_path):
    """解析toml文件"""
    try:
        import tomllib
    except ImportError:
        print("需要Python 3.11+或安装tomli")
        sys.exit(1)
    
    with open(file_path, 'rb') as f:
        data = tomllib.load(f)
    
    questions = []
    for q in data.get('questions', []):
        options = q.get('options', [])
        # 只取选项内容，不带A./B.前缀
        options_text = '|'.join([opt.get('text', '') for opt in options])
        questions.append({
            'number': q.get('number'),
            'type': q.get('type'),
            'stem': q.get('stem', ''),
            'options_raw': options_text,
            'answer': q.get('answer', ''),
            'notes': q.get('notes', '')
        })
    return questions


def clean_options(options_raw):
    """清理选项文本，去掉A./B./C./D.等前缀，只保留选项内容"""
    options = [opt.strip() for opt in options_raw.split('|')]
    cleaned = []
    for opt in options:
        clean = opt.strip()
        # 去掉选项前缀：A. B. C. D. E. F. 或 A、 B、 等
        for prefix in ['A.', 'B.', 'C.', 'D.', 'E.', 'F.', 'A、', 'B、', 'C、', 'D、', 'E、', 'F、']:
            if clean.startswith(prefix):
                clean = clean[len(prefix):].strip()
                break
        # 合并多余空格
        clean = ' '.join(clean.split())
        if clean:
            cleaned.append(clean)
    return cleaned


def format_question(q, cloze='{{c1::【】}}'):
    """格式化题目，添加cloze占位符"""
    stem = q['stem']
    q_type = q['type']
    
    # 处理各种括号格式
    placeholders = ['(    )', '(  )', '（    ）', '（）', '()']
    for ph in placeholders:
        if ph in stem:
            stem = stem.replace(ph, cloze)
            return stem
    
    # 判断题
    if q_type == '判断题':
        if stem.endswith('。'):
            stem = stem[:-1] + cloze + '。'
        else:
            stem = stem + cloze
    else:
        stem = stem + cloze
    
    return stem


def sanitize_tag(tag):
    """清理标签格式，确保每个标签以一级标签开头，层级内无空格"""
    if not tag:
        return tag
    
    # 按空格分割成多个标签
    tags = tag.strip().split()
    sanitized = []
    
    for t in tags:
        # 将层级内的空格替换为下划线（保留::作为层级分隔符）
        parts = t.split('::')
        parts = [p.replace(' ', '_') for p in parts]
        sanitized.append('::'.join(parts))
    
    return ' '.join(sanitized)


def convert_to_anki(questions, start_id=1, tag=''):
    """转换为Anki格式"""
    answer_map = {'A': 0, 'B': 1, 'C': 2, 'D': 3, 'E': 4, 'F': 5}
    
    lines = []
    current_id = start_id
    
    for q in questions:
        q_type = q['type']
        options = clean_options(q['options_raw'])
        answer_raw = q['answer']
        notes = q.get('notes', '')
        
        # 格式化题干
        stem = format_question(q)
        
        # 格式化选项
        if q_type == '判断题':
            options_text = '对||错'
            answer_num = '1' if answer_raw == '对' else '2'
        else:
            options_text = '||'.join(options)
            if q_type == '多选题':
                indices = [answer_map[ch] for ch in answer_raw if ch in answer_map]
                answer_num = '||'.join([str(i + 1) for i in indices])
            else:
                idx = answer_map.get(answer_raw, 0)
                answer_num = str(idx + 1)
        
        # 清理特殊字符
        stem = stem.replace('\t', ' ').replace('\n', ' ')
        options_text = options_text.replace('\t', ' ').replace('\n', ' ')
        notes = notes.replace('\t', ' ').replace('\n', ' ')
        
        # 字段顺序: id question options answer notes tags
        sanitized_tag = sanitize_tag(tag)
        lines.append(f"{current_id}\t{stem}\t{options_text}\t{answer_num}\t{notes}\t{sanitized_tag}")
        current_id += 1
    
    return lines, current_id


def get_max_id(file_path):
    """获取已有文件的最大ID"""
    if not os.path.exists(file_path):
        return 0
    
    max_id = 0
    with open(file_path, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if line and '\t' in line:
                try:
                    id_num = int(line.split('\t')[0])
                    max_id = max(max_id, id_num)
                except ValueError:
                    continue
    return max_id


def main():
    parser = argparse.ArgumentParser(description='Anki题目导入文件生成器')
    parser.add_argument('--input', required=True, help='输入文件路径')
    parser.add_argument('--output', required=True, help='输出文件路径')
    parser.add_argument('--start-id', type=int, default=1, help='起始ID（默认1）')
    parser.add_argument('--tag', default='', help='标签（默认使用章节名）')
    parser.add_argument('--append', action='store_true', help='追加到已有文件')
    
    args = parser.parse_args()
    
    # 解析输入文件
    ext = os.path.splitext(args.input)[1].lower()
    if ext == '.xlsx':
        questions = parse_xlsx(args.input)
    elif ext == '.toml':
        questions = parse_toml(args.input)
    else:
        print(f"暂不支持的格式: {ext}")
        sys.exit(1)
    
    # 确定起始ID
    start_id = args.start_id
    if args.append and os.path.exists(args.output):
        start_id = get_max_id(args.output) + 1
    
    # 转换格式
    lines, next_id = convert_to_anki(questions, start_id, args.tag)
    
    # 写入文件
    mode = 'a' if args.append and os.path.exists(args.output) else 'w'
    with open(args.output, mode, encoding='utf-8') as f:
        if mode == 'a' and not args.append:
            f.write('\n')
        f.write('\n'.join(lines))
        if not args.append:
            f.write('\n')
    
    print(f"转换完成: {len(questions)}题")
    print(f"ID范围: {start_id}-{next_id - 1}")
    print(f"输出文件: {args.output}")
    print(f"\n导入Anki时请选择 AwesomeSelect-3.x 模板")


if __name__ == '__main__':
    main()
