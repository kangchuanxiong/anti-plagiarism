#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
增强版反抄袭检测工具（仅使用标准库，不修改原有 anti_plagiarism_tool.py）

在原有“精确段落匹配”思路基础上的四项增强：
  模块A  文本归一化 + 模板过滤后的精确匹配
         （全角转半角、去空白标点、忽略大小写；全班高频段落自动识别为模板并剔除，
           也可通过 --template 指定白名单文件，每行一个模板段落）
  模块B  N-gram 指纹检测（Shingling + Winnowing，MOSS 同类算法）
         （可捕捉局部抄袭与少量改写；全班共享指纹自动降权，避免模板误报）
  模块C  句级相似度匹配（difflib，相似度 >= 阈值记为重复，可精确定位抄袭句子）
  模块D  docx 修订会话 ID（w:rsid）同源检测
         （从同一原始文档复制/另存的文件会残留大量相同 rsid，学生难以伪造）

用法:
    python anti_plagiarism_enhanced.py [文件夹路径]
    不带参数运行时将提示输入文件夹路径。

输出（CSV，均为 utf-8-sig，Excel 可直接打开）:
    enhanced_template_paragraphs.csv   识别出的模板段落清单
    enhanced_exact_match.csv           模块A 归一化精确匹配结果
    enhanced_fingerprint.csv           模块B 指纹检测结果
    enhanced_sentence_match.csv        模块C 句级匹配结果（含证据句样例）
    enhanced_rsid_analysis.csv         模块D rsid 同源分析结果
"""

import argparse
import csv
import hashlib
import os
import re
import sys
import unicodedata
import xml.etree.ElementTree as ET
import zipfile
from collections import Counter, defaultdict

# ========================
# 可调参数
# ========================

# 模块B：指纹参数
K_GRAM = 10                 # n-gram 长度（字符）
WINNOW_WINDOW = 12          # winnowing 滑窗大小
COMMON_FINGERPRINT_RATIO = 0.5   # 出现在超过该比例文档中的指纹视为公共指纹（模板/常见语料）

# 模块C：句级匹配参数
SENT_SIM_THRESHOLD = 0.90   # 句子相似度阈值（difflib ratio）
MIN_SENT_LEN = 8            # 参与句级匹配的最短归一化句长（字符）
MAX_CANDIDATES = 60         # 每个句子最多比对的候选句数量（按共享指纹数排序取前 N）

# 模板自动识别参数：段落（归一化后）出现在 >= 该比例的文档中即视为模板
TEMPLATE_FREQ_RATIO = 0.5
# 模板段落最小长度：太短的归一化段落（如“摘要”）即使全班出现也不作为抄袭证据
MIN_TEMPLATE_LEN = 4

# 模块D：rsid 参数
RSID_MIN_SHARED = 20        # 两文档至少共享的 rsid 数量
RSID_OVERLAP_COEF = 0.30    # 重叠系数阈值 = 共同rsid数 / min(两文档rsid总数)

DOC_NS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


# ========================
# docx 读取
# ========================

def extract_docx_data(docx_path):
    """从 docx 中提取段落文本与 rsid 集合。失败返回 None。"""
    try:
        with zipfile.ZipFile(docx_path, 'r') as zf:
            if 'word/document.xml' not in zf.namelist():
                return None
            with zf.open('word/document.xml') as f:
                root = ET.parse(f).getroot()

            # 提取段落文本（含表格、文本框内的段落）
            paragraphs = []
            for para in root.iter(DOC_NS + 'p'):
                text = ''.join(t.text or '' for t in para.iter(DOC_NS + 't')).strip()
                if text:
                    paragraphs.append(text)

            # 提取修订会话 ID（w:rsid* 属性值）
            rsids = set()
            for elem in root.iter():
                for key, value in elem.attrib.items():
                    if 'rsid' in key.lower() and value:
                        rsids.add(value)

            return {'paragraphs': paragraphs, 'rsids': rsids}
    except Exception as e:
        print(f"  [错误] 处理 {os.path.basename(docx_path)} 失败: {e}")
        return None


# ========================
# 归一化
# ========================

_PUNCT_RE = re.compile(r'[^\w]+')


def normalize_text(text):
    """归一化：全角转半角、转小写、去空白与标点，仅保留字母/数字/下划线/CJK等词字符。"""
    text = unicodedata.normalize('NFKC', text).lower()
    return _PUNCT_RE.sub('', text)


def split_sentences(paragraphs):
    """把段落切分为句子。"""
    sentences = []
    for para in paragraphs:
        for sent in re.split(r'[。！？!?；;]+|[\r\n\t]+', para):
            sent = sent.strip()
            if len(sent) >= 6:
                sentences.append(sent)
    return sentences


# ========================
# 模板识别与过滤
# ========================

def load_template_file(path):
    """读取用户提供的模板白名单文件（每行一个模板段落）。"""
    if not path:
        return set()
    try:
        with open(path, 'r', encoding='utf-8-sig') as f:
            items = {normalize_text(line) for line in f if len(normalize_text(line)) >= MIN_TEMPLATE_LEN}
        if items:
            print(f"已从 {path} 载入 {len(items)} 条模板段落（白名单）")
        return items
    except Exception as e:
        print(f"[警告] 模板文件读取失败: {e}")
        return set()


def detect_template_paragraphs(all_norm_paras, doc_count):
    """自动识别模板段落：出现在足够多文档中的高频归一化段落。"""
    para_doc_freq = Counter()
    for norm_paras in all_norm_paras.values():
        for p in set(norm_paras):
            para_doc_freq[p] += 1

    min_docs = max(2, int(doc_count * TEMPLATE_FREQ_RATIO))
    templates = {p for p, cnt in para_doc_freq.items()
                 if cnt >= min_docs and len(p) >= MIN_TEMPLATE_LEN}
    return templates, para_doc_freq


def filter_template(items, template_set):
    """过滤掉模板段落/句子。"""
    return [x for x in items if x not in template_set]


# ========================
# 模块A：归一化精确匹配
# ========================

def run_exact_match(doc_data, template_set):
    """逐对比较归一化后（剔除模板）的段落集合，输出每篇文档的最相似文档与重复率。"""
    results = []
    names = sorted(doc_data)
    for target in names:
        t_set = set(filter_template(doc_data[target]['norm_paras'], template_set))
        t_chars = sum(len(p) for p in t_set)
        best_other, best_rate = '', 0.0
        for other in names:
            if other == target:
                continue
            o_set = set(doc_data[other]['norm_paras'])
            common_chars = sum(len(p) for p in t_set & o_set)
            rate = common_chars / t_chars if t_chars else 0.0
            if rate > best_rate:
                best_other, best_rate = other, rate
        if t_chars == 0:
            print(f"  {target}: 剔除模板后无剩余段落，跳过")
            continue
        results.append({
            '文档': target,
            '最相似文档': best_other,
            '归一化精确重复率': round(best_rate * 100, 2),
            '非模板段落总字数': t_chars,
        })
        print(f"  {target} -> {best_other}: {best_rate * 100:.2f}%")
    return results


# ========================
# 模块B：N-gram 指纹（Winnowing）
# ========================

def _gram_hash(gram):
    return int.from_bytes(hashlib.md5(gram.encode('utf-8')).digest()[:8], 'big')


def winnow_fingerprints(text, k=K_GRAM, w=WINNOW_WINDOW):
    """Winnowing 指纹：k-gram 哈希后按窗口取最小值。返回指纹集合。"""
    hashes = [_gram_hash(text[i:i + k]) for i in range(len(text) - k + 1)]
    if not hashes:
        return set()
    fingerprints = set()
    for start in range(0, max(1, len(hashes) - w + 1)):
        window = hashes[start:start + w]
        # 取窗口内最小值（相同取最右出现，保证覆盖性）
        fingerprints.add(min(reversed(window)))
    return fingerprints


def run_fingerprint_match(doc_data, template_set):
    """基于指纹集合的 Jaccard/覆盖率比较，公共指纹（全班高频）自动剔除。"""
    fps = {}
    for name, data in doc_data.items():
        body_text = ''.join(filter_template(data['norm_paras'], template_set))
        fps[name] = winnow_fingerprints(body_text)

    # 全班共享指纹降权
    n_docs = len(fps)
    df = Counter()
    for f in fps.values():
        for h in f:
            df[h] += 1
    min_docs = max(2, int(n_docs * COMMON_FINGERPRINT_RATIO))
    common_fps = {h for h, cnt in df.items() if cnt >= min_docs}

    results = []
    names = sorted(fps)
    for target in names:
        t_fps = fps[target] - common_fps
        best_other, best_rate = '', 0.0
        for other in names:
            if other == target:
                continue
            o_fps = fps[other] - common_fps
            if not t_fps:
                continue
            shared = len(t_fps & o_fps)
            rate = shared / len(t_fps)
            if rate > best_rate:
                best_other, best_rate = other, rate
        results.append({
            '文档': target,
            '最相似文档': best_other,
            '指纹重合率': round(best_rate * 100, 2),
            '有效指纹数': len(t_fps),
            '剔除的公共指纹数': len(fps[target] & common_fps),
        })
        print(f"  {target} -> {best_other}: {best_rate * 100:.2f}% "
              f"(有效指纹 {len(t_fps)})")
    return results


# ========================
# 模块C：句级相似度匹配
# ========================

def build_sentence_index(doc_data, template_set):
    """构建句子级倒排索引：8-gram 哈希 -> {(文档, 句下标)}，短句用整句哈希。"""
    index = defaultdict(set)
    short_hash = defaultdict(set)
    for name, data in doc_data.items():
        for i, sent in enumerate(filter_template(data['norm_sents'], template_set)):
            if len(sent) < MIN_SENT_LEN:
                continue
            if len(sent) < K_GRAM + 2:
                short_hash[_gram_hash(sent)].add((name, i))
            else:
                for j in range(len(sent) - K_GRAM + 1):
                    index[_gram_hash(sent[j:j + K_GRAM])].add((name, i))
    return index, short_hash


def run_sentence_match(doc_data, template_set):
    """逐句比对：目标句与候选句 difflib ratio >= 阈值即记为重复（每句只记一次，取最佳配对）。"""
    from difflib import SequenceMatcher

    index, short_hash = build_sentence_index(doc_data, template_set)
    results = []
    names = sorted(doc_data)

    for target in names:
        sent_pairs = [(norm, raw)
                      for norm, raw in zip(doc_data[target]['norm_sents'],
                                           doc_data[target]['sents'])
                      if norm not in template_set]
        target_sents = [norm for norm, _ in sent_pairs]
        raw_sents = [raw for _, raw in sent_pairs]

        total_chars = sum(len(s) for s in target_sents)
        matched_chars = 0
        matched_count = 0
        evidence = []   # (相似度, 证据句原文, 最相似文档)

        for i, sent in enumerate(target_sents):
            # 收集候选句（排除自己文档、排除模板句）
            if len(sent) < K_GRAM + 2:
                candidates = short_hash.get(_gram_hash(sent), set())
            else:
                gram_hashes = {_gram_hash(sent[j:j + K_GRAM])
                               for j in range(len(sent) - K_GRAM + 1)}
                counter = Counter()
                for h in gram_hashes:
                    for (doc, idx) in index.get(h, ()):
                        if doc != target:
                            counter[(doc, idx)] += 1
                candidates = {pair for pair, _ in
                              counter.most_common(MAX_CANDIDATES)}
            if not candidates:
                continue

            best_ratio, best_pair = 0.0, None
            for (doc, idx) in candidates:
                other_sent = filter_template(doc_data[doc]['norm_sents'], template_set)
                if idx >= len(other_sent):
                    continue
                ratio = SequenceMatcher(None, sent, other_sent[idx]).ratio()
                if ratio >= SENT_SIM_THRESHOLD and ratio > best_ratio:
                    best_ratio, best_pair = ratio, (doc, idx)

            if best_pair:
                matched_chars += len(sent)
                matched_count += 1
                if len(evidence) < 5:
                    evidence.append((best_ratio, raw_sents[i], best_pair[0]))

        rate = matched_chars / total_chars if total_chars else 0.0
        results.append({
            '文档': target,
            '疑似重复句数': matched_count,
            '有效句总字数': total_chars,
            '句级重复率': round(rate * 100, 2),
            '证据句样例': ' | '.join(
                f"[{r:.2f}]{s[:40]}..." if len(s) > 40 else f"[{r:.2f}]{s}"
                for r, s, _ in sorted(evidence, reverse=True)
            ),
        })
        print(f"  {target}: 重复率 {rate * 100:.2f}% ({matched_count} 句)")

    return results


# ========================
# 模块D：rsid 同源检测
# ========================

def run_rsid_analysis(doc_data):
    """两两比较 rsid 集合，重叠系数超过阈值即报告“疑似同源”。"""
    results = []
    names = sorted(doc_data)
    pairs_found = 0
    for i, a in enumerate(names):
        for b in names[i + 1:]:
            set_a, set_b = doc_data[a]['rsids'], doc_data[b]['rsids']
            if not set_a or not set_b:
                continue
            shared = len(set_a & set_b)
            coef = shared / min(len(set_a), len(set_b))
            if shared >= RSID_MIN_SHARED and coef >= RSID_OVERLAP_COEF:
                pairs_found += 1
                verdict = '高度疑似同源' if coef >= 0.6 else '疑似同源'
                results.append({
                    '文档A': a, '文档B': b,
                    '共享rsid数': shared,
                    'A的rsid数': len(set_a), 'B的rsid数': len(set_b),
                    '重叠系数': round(coef, 3),
                    '判定': verdict,
                })
                print(f"  {a} <-> {b}: 共享 {shared} 个 rsid "
                      f"(系数 {coef:.2f}) -> {verdict}")
    if not pairs_found:
        print("  未发现疑似同源文档")
    return results


# ========================
# CSV 输出
# ========================

def save_csv(path, rows, fieldnames):
    try:
        with open(path, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(rows)
        print(f"已保存: {path}")
    except Exception as e:
        print(f"[错误] 保存 {path} 失败: {e}")


# ========================
# 主流程
# ========================

def load_documents(folder):
    doc_data = {}
    files = sorted(f for f in os.listdir(folder)
                   if os.path.isfile(os.path.join(folder, f)))
    print(f"共发现 {len(files)} 个文件，开始解析...\n")
    for name in files:
        data = extract_docx_data(os.path.join(folder, name))
        if not data or not data['paragraphs']:
            if data is not None:
                print(f"  [跳过] {name}: 无有效段落（可能不是 docx）")
            continue
        data['sents'] = split_sentences(data['paragraphs'])
        data['norm_paras'] = [normalize_text(p) for p in data['paragraphs']]
        data['norm_paras'] = [p for p in data['norm_paras'] if p]
        data['norm_sents'] = [normalize_text(s) for s in data['sents']]
        doc_data[name] = data
    print(f"\n成功解析 {len(doc_data)} 个文档\n")
    return doc_data


def run_enhanced_detection(folder, template_path=None):
    """运行增强检测并生成各模块的 CSV 报告。"""
    print("=" * 60)
    print("增强版反抄袭检测工具")
    print("模块A 归一化精确匹配 | 模块B 指纹检测 | 模块C 句级匹配 | 模块D rsid同源")
    print("=" * 60)

    doc_data = load_documents(folder)
    if len(doc_data) < 2:
        print("有效文档不足 2 个，无法进行对比，程序退出。")
        return False

    print("\n【步骤0】模板段落识别与过滤")
    print("-" * 50)
    template_set = load_template_file(template_path)
    auto_templates, freq = detect_template_paragraphs(
        {k: v['norm_paras'] for k, v in doc_data.items()}, len(doc_data))
    template_set |= auto_templates
    print(f"自动识别模板段落 {len(auto_templates)} 条，"
          f"模板总计 {len(template_set)} 条")
    save_csv('enhanced_template_paragraphs.csv',
             [{'模板段落': p[:100], '出现文档数': freq.get(p, 0)}
              for p in sorted(template_set)],
             ['模板段落', '出现文档数'])

    print("\n【模块A】归一化精确匹配（已剔除模板）")
    print("-" * 50)
    exact_rows = run_exact_match(doc_data, template_set)
    save_csv('enhanced_exact_match.csv', exact_rows,
             ['文档', '最相似文档', '归一化精确重复率', '非模板段落总字数'])

    print("\n【模块B】N-gram 指纹检测（Winnowing）")
    print("-" * 50)
    fp_rows = run_fingerprint_match(doc_data, template_set)
    save_csv('enhanced_fingerprint.csv', fp_rows,
             ['文档', '最相似文档', '指纹重合率', '有效指纹数', '剔除的公共指纹数'])

    print("\n【模块C】句级相似度匹配")
    print("-" * 50)
    sent_rows = run_sentence_match(doc_data, template_set)
    save_csv('enhanced_sentence_match.csv', sent_rows,
             ['文档', '疑似重复句数', '有效句总字数', '句级重复率', '证据句样例'])

    print("\n【模块D】docx rsid 同源检测")
    print("-" * 50)
    rsid_rows = run_rsid_analysis(doc_data)
    save_csv('enhanced_rsid_analysis.csv', rsid_rows,
             ['文档A', '文档B', '共享rsid数', 'A的rsid数', 'B的rsid数',
              '重叠系数', '判定'])

    print("\n" + "=" * 60)
    print("全部检测完成！")
    print("判读建议：模块C 句级重复率与模块D rsid 判定证据最直观，")
    print("模块A/B 适合快速筛查；多模块同时命中高分的文档对优先人工复核。")
    print("=" * 60)
    return True


def main():
    parser = argparse.ArgumentParser(description='增强版反抄袭检测工具（仅标准库）')
    parser.add_argument('folder', nargs='?', help='包含 docx 文件的文件夹路径')
    parser.add_argument('--template', help='模板白名单文件（每行一个模板段落，可选）')
    args = parser.parse_args()

    folder = args.folder or input("请输入文件夹路径: ").strip()
    if not os.path.isdir(folder):
        print(f"错误: 目录 {folder} 不存在")
        sys.exit(1)

    if not run_enhanced_detection(folder, args.template):
        sys.exit(1)


if __name__ == '__main__':
    main()
