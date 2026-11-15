#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
综合反抄袭检测工具
按顺序执行：文档元数据分析、文字抄袭检测、图片抄袭检测
"""

import csv
import hashlib
import os
import shutil
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime
from collections import defaultdict

# 只保留必要的XML命名空间
NAMESPACES = {
    'cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
    'dc': 'http://purl.org/dc/elements/1.1/',
    'dcterms': 'http://purl.org/dc/terms/'
}

# ========================
# 文档元数据分析相关函数
# ========================

def parse_iso_datetime(iso_str):
    """将ISO格式时间（如2023-01-01T12:00:00Z）转换为本地时间字符串，支持多种格式"""
    try:
        # 处理常见的时间格式变体
        if iso_str:
            # 处理带Z的UTC时间
            if iso_str.endswith('Z'):
                iso_str = iso_str[:-1] + '+00:00'
            # 处理没有时区信息的格式
            elif 'T' in iso_str and not any(tz in iso_str for tz in ['+', '-', 'Z']):
                # 假设为UTC时间
                iso_str = iso_str + '+00:00'
            
            dt = datetime.fromisoformat(iso_str)
            return dt.astimezone().strftime('%Y-%m-%d %H:%M:%S')  # 转换为本地时间
        return ''
    except (ValueError, TypeError):
        return ''

def parse_core_properties(core_xml_content):
    """解析core.xml获取核心属性，只提取四个必要信息"""
    try:
        root = ET.fromstring(core_xml_content)
        
        # 提取作者信息
        author = root.findtext('dc:creator', '', namespaces=NAMESPACES)
        if not author:
            for elem in root.iter():
                if elem.tag.endswith('creator') or elem.tag == 'creator':
                    author = elem.text or ''
                    break
        
        # 提取上次修改者信息
        last_modified_by = root.findtext('cp:lastModifiedBy', '', namespaces=NAMESPACES)
        if not last_modified_by:
            for elem in root.iter():
                if elem.tag.endswith('lastModifiedBy') or elem.tag == 'lastModifiedBy':
                    last_modified_by = elem.text or ''
                    break
        
        # 解析创建时间
        created_str = root.findtext('dcterms:created', '', namespaces=NAMESPACES)
        if not created_str:
            for elem in root.iter():
                if elem.tag.endswith('created') or elem.tag == 'created':
                    created_str = elem.text or ''
                    break
        
        # 解析修改时间
        modified_str = root.findtext('dcterms:modified', '', namespaces=NAMESPACES)
        if not modified_str:
            for elem in root.iter():
                if elem.tag.endswith('modified') or elem.tag == 'modified':
                    modified_str = elem.text or ''
                    break
        
        created_time = parse_iso_datetime(created_str) if created_str else ''
        modified_time = parse_iso_datetime(modified_str) if modified_str else ''
        
        return {
            '作者': author,
            '上次修改者': last_modified_by,
            '创建时间': created_time,
            '上次修改时间': modified_time
        }
    except Exception:
        return {
            '作者': '',
            '上次修改者': '',
            '创建时间': '',
            '上次修改时间': ''
        }

def get_docx_properties(file_path):
    """获取单个文件的四个必要属性信息"""
    try:
        if not os.path.exists(file_path) or not os.path.isfile(file_path):
            print(f"无法处理: {os.path.basename(file_path)}")
            return None
        
        # 尝试作为zip文件打开
        with zipfile.ZipFile(file_path, 'r') as zf:
            # 尝试查找core.xml
            core_xml_paths = []
            for name in zf.namelist():
                if name.lower().endswith('core.xml'):
                    core_xml_paths.append(name)
            
            if not core_xml_paths:
                print(f"无法处理: {os.path.basename(file_path)}")
                return None
                
            # 解析核心属性
            core_props = None
            for core_path in core_xml_paths:
                try:
                    with zf.open(core_path) as f:
                        core_content = f.read().decode('utf-8')
                    core_props = parse_core_properties(core_content)
                    break
                except Exception:
                    continue
            
            if core_props is None:
                print(f"无法处理: {os.path.basename(file_path)}")
                return None
        
        # 只返回四个必要属性
        return {
            '文件名': os.path.basename(file_path),
            '上次修改时间': core_props['上次修改时间'],
            '创建时间': core_props['创建时间'],
            '作者': core_props['作者'],
            '上次修改者': core_props['上次修改者']
        }
        
    except Exception:
        print(f"无法处理: {os.path.basename(file_path)}")
        return None

def process_all_files(folder):
    """处理文件夹中的所有文件"""
    results = []
    for filename in os.listdir(folder):
        # 使用 os.path.join 构建完整路径
        file_path = os.path.join(folder, filename)
        if os.path.isfile(file_path):
            props = get_docx_properties(file_path)
            if props:
                results.append(props)
    return results

def save_metadata_to_csv(data, output_file="plagiarism_metadata_analysis.csv"):
    """将元数据结果保存为CSV文件"""
    try:
        if not data:
            return
            
        with open(output_file, 'w', newline='', encoding='utf-8-sig') as f:
            # 使用数据中的第一个元素的键作为表头
            if data:
                writer = csv.DictWriter(f, fieldnames=data[0].keys())
                writer.writeheader()
                writer.writerows(data)
            
        print(f"元数据信息已成功保存到CSV文件: {output_file}")
    except Exception:
        pass

def find_similar_files(docx_info):
    """查找具有相同作者和创建时间的文件"""
    similar_groups = {}
    
    for doc_info in docx_info:
        author = doc_info.get('作者', '')
        created_time = doc_info.get('创建时间', '')
        
        if not author or not created_time:
            continue
        
        key = (author, created_time)
        if key not in similar_groups:
            similar_groups[key] = []
        similar_groups[key].append(doc_info)
    
    result = {k: v for k, v in similar_groups.items() if len(v) > 1}
    return result

def display_similar_files(similar_files):
    """显示具有相同作者和创建时间的文件组"""
    if not similar_files:
        print("没有找到作者和创建时间相同的文件")
        return
    
    print(f"\n找到 {len(similar_files)} 组作者和创建时间相同的文件：\n")
    for idx, (key, files) in enumerate(similar_files.items(), 1):
        author, created_time = key
        print(f"组 {idx}: 作者 '{author}', 创建时间 '{created_time}'")
        print(f"  包含 {len(files)} 个文件:")
        for file in files:
            print(f"    - {file['文件名']}")
        print()

def analyze_document_metadata(docx_directory):
    """执行文档元数据分析"""
    print("\n【第一部分】文档元数据分析")
    print("=" * 50)
    
    if not os.path.isdir(docx_directory):
        print(f"错误: 目录 {docx_directory} 不存在")
        return False
    
    # 直接处理所有文件
    processed_files = process_all_files(docx_directory)
    print(f"处理完成，成功处理 {len(processed_files)} 个文件")
    save_metadata_to_csv(processed_files)
    
    if processed_files:
        # 查找并显示相似文件
        similar_files = find_similar_files(processed_files)
        display_similar_files(similar_files)
        return True
    else:
        print("未找到可正常处理的文件")
        return False

# ========================
# 文字抄袭检测相关函数
# ========================

def extract_paragraphs_from_docx(docx_path):
    """
    从docx文件中提取所有段落文本
    """
    paragraphs = []
    
    try:
        # 检查文件是否存在
        if not os.path.exists(docx_path):
            print(f"错误：文件 {docx_path} 不存在")
            return []
            
        # 检查文件是否为docx格式（基本检查）
        if not docx_path.lower().endswith('.docx'):
            print(f"错误：文件 {docx_path} 不是docx格式")
            return []
            
        # 以zip格式打开docx文件
        with zipfile.ZipFile(docx_path, 'r') as zip_ref:
            # 检查是否是有效的docx文件
            if 'word/document.xml' not in zip_ref.namelist():
                print(f"错误：文件 {docx_path} 不是有效的docx文件，缺少必要的结构")
                return []
                
            # 读取主要的文档内容XML
            with zip_ref.open('word/document.xml') as f:
                # 解析XML
                tree = ET.parse(f)
                root = tree.getroot()
                
                # Word XML使用的命名空间
                ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
                
                # 查找所有段落元素
                para_elements = root.findall('.//w:p', namespaces=ns)
                
                for para in para_elements:
                    # 提取段落中的文本
                    text_parts = []
                    # 查找段落中的所有文本元素
                    text_elements = para.findall('.//w:t', namespaces=ns)
                    
                    for text_elem in text_elements:
                        if text_elem.text:
                            text_parts.append(text_elem.text)
                    
                    # 合并段落的所有文本部分
                    paragraph_text = ''.join(text_parts).strip()
                    
                    # 只添加非空段落
                    if paragraph_text:
                        paragraphs.append(paragraph_text)
        
        return paragraphs
        
    except zipfile.BadZipFile:
        print(f"错误：文件 {docx_path} 不是有效的zip文件（docx格式基于zip）")
        return []
    except ET.ParseError as e:
        print(f"错误：解析文件 {docx_path} 的XML内容时出错: {str(e)}")
        return []
    except PermissionError:
        print(f"错误：没有权限访问文件 {docx_path}")
        return []
    except Exception as e:
        print(f"处理文件 {docx_path} 时出错: {str(e)}")
        return []

def calculate_similarity_rate_by_word_count(paragraphs1, paragraphs2):
    """
    计算两个文档段落的双向相似度
    返回相似段落总字数在两个文档各自总字数中占比
    """
    if not paragraphs1 or not paragraphs2:
        return (0.0, 0.0)
    
    # 计算交集（相似的段落）
    set1 = set(paragraphs1)
    set2 = set(paragraphs2)
    common_paragraphs = set1.intersection(set2)
    
    # 计算相似段落的总字数
    common_words_count = sum(len(para) for para in common_paragraphs)
    
    # 计算两个文档的总字数
    words_count1 = sum(len(para) for para in set1)
    words_count2 = sum(len(para) for para in set2)
    
    # 计算双向相似度
    similarity_rate_1_to_2 = common_words_count / words_count1 if words_count1 > 0 else 0.0
    similarity_rate_2_to_1 = common_words_count / words_count2 if words_count2 > 0 else 0.0
    
    return (similarity_rate_1_to_2, similarity_rate_2_to_1)

def load_all_document_paragraphs(docx_dir):
    """
    加载目录中所有文档的段落
    返回包含所有成功处理文档段落的字典
    """
    # 存储每个文档的段落
    doc_paragraphs = {}
    failed_documents = []
    
    # 获取目录下所有文件（不再筛选docx格式）
    all_files = []
    for f in os.listdir(docx_dir):
        file_path = os.path.join(docx_dir, f)
        if os.path.isfile(file_path):
            all_files.append(f)
    
    if not all_files:
        print("未找到任何文件")
        return {}
    
    print(f"找到 {len(all_files)} 个文件，正在提取段落...")
    
    # 提取每个文档的段落
    for file in all_files:
        # 使用 os.path.join 构建完整路径
        file_path = os.path.join(docx_dir, file)
        paragraphs = extract_paragraphs_from_docx(file_path)
        
        if paragraphs:
            doc_paragraphs[file] = paragraphs
            print(f"从 {file} 中提取了 {len(paragraphs)} 个段落")
        else:
            failed_documents.append(file)
            # 注意：extract_paragraphs_from_docx函数已经输出了具体的错误信息
    
    # 如果有处理失败的文档，显示出来
    if failed_documents:
        print(f"\n处理失败的文档 ({len(failed_documents)} 个):")
        for doc in failed_documents:
            print(f"  - {doc}")
    
    return doc_paragraphs

def compare_target_vs_others(doc_paragraphs):
    """
    比较每个目标文件与其他所有文件的相似度
    返回每个文件最相似的文档名称和重复率
    """
    # 如果没有提取到任何段落，直接返回
    if not doc_paragraphs:
        return {}
    
    # 比较每个文档与其他所有文档
    results = defaultdict(dict)
    filenames = list(doc_paragraphs.keys())
    total_comparisons = len(filenames) * (len(filenames) - 1)
    
    print(f"正在进行 {total_comparisons} 组目标文件与其他文件的对比...\n")
    
    for target_doc in filenames:
        # 初始化最相似文档信息
        most_similar_doc = ""
        highest_similarity = 0.0
        
        for other_doc in filenames:
            if target_doc == other_doc:
                continue  # 跳过与自己比较
                
            try:
                # 计算相似度
                target_to_other_similarity, _ = calculate_similarity_rate_by_word_count(
                    doc_paragraphs[target_doc], 
                    doc_paragraphs[other_doc]
                )
                
                # 转换为百分比
                similarity_percent = round(target_to_other_similarity * 100, 2)
                
                # 更新最相似文档信息
                if similarity_percent > highest_similarity:
                    highest_similarity = similarity_percent
                    most_similar_doc = other_doc
                
            except Exception as e:
                print(f"\n比较文档 {target_doc} 和 {other_doc} 时出错: {str(e)}")
        
        # 存储当前目标文件的结果
        results[target_doc] = {
            'most_similar_doc': most_similar_doc,
            'highest_similarity': highest_similarity
        }
        
        # 输出当前目标文档与最相似文档的重复率
        print(f"目标文档 '{target_doc}' 与最相似文档 '{most_similar_doc}' 的重复率: {highest_similarity:.2f}%")
    
    print("\n目标文件与其他文件对比完成")
    return results

def compare_documents_to_collection(doc_paragraphs):
    """
    比较每个文档与其他所有文档的段落集合的相似度
    返回包含相似度计算结果的字典
    """
    # 如果没有提取到任何段落，直接返回
    if not doc_paragraphs:
        return {}
    
    # 比较每个文档与其他所有文档的集合
    results = defaultdict(dict)
    filenames = list(doc_paragraphs.keys())
    
    print(f"正在比较每个文档与其他文档集合的相似度 ({len(filenames)} 个文档)...")
    
    for i, target_doc in enumerate(filenames):
        try:
            target_paragraphs = doc_paragraphs[target_doc]
            
            # 收集其他所有文档的段落
            other_paragraphs = []
            for j, other_doc in enumerate(filenames):
                if i != j:
                    other_paragraphs.extend(doc_paragraphs[other_doc])
            
            # 去重，确保每个段落只出现一次
            other_paragraphs_set = set(other_paragraphs)
            
            # 计算相似度（只取目标文档到集合的相似度）
            target_to_collection_similarity, _ = calculate_similarity_rate_by_word_count(
                target_paragraphs, 
                other_paragraphs_set
            )
            
            # 转换为百分比
            target_to_collection_similarity_percent = round(target_to_collection_similarity * 100, 2)
            
            # 存储结果
            results[target_doc] = {
                'target_to_collection_similarity_rate': target_to_collection_similarity_percent,
                'other_docs_count': len(filenames) - 1,
                'other_paragraphs_count': len(other_paragraphs_set)
            }
            
            # 简化输出格式
            print(f"{target_doc} 与其他文档集合的相似度: {target_to_collection_similarity_percent}%")
            
        except Exception as e:
            print(f"比较文档 {target_doc} 与其他文档集合时出错: {str(e)}")
    
    return results

def detect_text_plagiarism(docx_directory):
    """执行文字抄袭检测"""
    print("\n【第二部分】文档内容相似度比较（文字抄袭检测）")
    print("=" * 50)
    
    if not os.path.isdir(docx_directory):
        print(f"错误: 目录 {docx_directory} 不存在")
        return False
    
    # 一次性提取所有文档的段落，避免重复操作
    doc_paragraphs = load_all_document_paragraphs(docx_directory)
    
    if not doc_paragraphs:
        print("无法提取文档段落，程序退出")
        return False
    
    # 执行目标文件与其他文件逐个对比
    print("\n开始比较每个目标文件与其他所有文件的相似度...")
    target_results = compare_target_vs_others(doc_paragraphs)
    
    # 执行文档与集合比较
    print("\n开始比较每个文档与其他文档集合的相似度...")
    collection_results = compare_documents_to_collection(doc_paragraphs)
    
    # 将两个结果合并为一个CSV文件
    if target_results and collection_results:
        # 合并后的输出文件
        merged_output = "plagiarism_combined_results.csv"
        try:
            with open(merged_output, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                # 写入表头
                writer.writerow([
                    "目标文档", 
                    "最相似文档", 
                    "目标文档与最相似文档重复率", 
                    "目标文档与其他文档集合重复率"
                ])
                
                # 写入所有合并结果
                for target_doc in sorted(target_results.keys()):
                    # 获取目标文档与最相似文档的信息
                    target_info = target_results[target_doc]
                    most_similar_doc = target_info['most_similar_doc']
                    highest_similarity = target_info['highest_similarity']
                    
                    # 获取目标文档与其他文档集合的信息
                    collection_info = collection_results.get(target_doc, {})
                    collection_similarity = collection_info.get('target_to_collection_similarity_rate', 0.0)
                    
                    # 写入一行数据
                    writer.writerow([
                        target_doc, 
                        most_similar_doc, 
                        highest_similarity, 
                        collection_similarity
                    ])
            
            print(f"合并结果已保存到 {merged_output}")
            return True
        except Exception as e:
            print(f"保存合并结果到CSV文件时出错: {str(e)}")
            return False
    else:
        print("无法获取足够的相似度数据来生成合并结果")
        return False

# ========================
# 图片抄袭检测相关函数
# ========================

def extract_images_from_docx(docx_path, extract_dir):
    """
    从docx文件中提取图片并保存到指定目录
    返回元组：(提取的图片列表, 是否成功)
    """
    try:
        os.makedirs(extract_dir, exist_ok=True)
        
        with zipfile.ZipFile(docx_path, 'r') as zip_ref:
            image_files = [
                f for f in zip_ref.namelist() 
                if f.startswith('word/media/') or 
                   ('media/' in f and not f.endswith('/') and 
                    any(f.lower().endswith(ext) for ext in ['.jpg', '.jpeg', '.png', '.gif', '.bmp']))
            ]
            
            extracted_files = []
            for img_file in set(image_files):  # 去重
                try:
                    img_name = os.path.basename(img_file)
                    save_path = os.path.join(extract_dir, img_name)
                    
                    with open(save_path, 'wb') as f:
                        f.write(zip_ref.read(img_file))
                    extracted_files.append(save_path)
                except Exception:
                    continue
                    
            return extracted_files, True
            
    except Exception:
        return [], False

def calculate_file_hash(file_path, hash_algorithm='md5'):
    """
    计算文件的哈希值
    """
    try:
        hash_obj = hashlib.new(hash_algorithm)
        
        with open(file_path, 'rb') as f:
            # 分块读取文件（4KB每块）
            while chunk := f.read(4096):
                hash_obj.update(chunk)
        
        return hash_obj.hexdigest()
    except Exception:
        return None

def find_duplicate_images(docx_dir, temp_dir=None):
    """
    查找指定目录下所有文件中的重复图片
    """
    # 如果未指定临时目录，使用相对路径下的'temp_word_images'
    if temp_dir is None:
        temp_dir = 'temp_word_images'
    
    hash_to_images = defaultdict(list)
    
    try:
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)
        os.makedirs(temp_dir, exist_ok=True)
    except Exception as e:
        print(f"创建临时目录失败: {str(e)}")
        raise
    
    try:
        # 获取目录下所有文件（不筛选扩展名）
        all_files = []
        for f in os.listdir(docx_dir):
            file_path = os.path.join(docx_dir, f)
            if os.path.isfile(file_path):
                all_files.append(f)
        
        if not all_files:
            print("当前目录下没有文件")
            return {}
        
        print(f"找到 {len(all_files)} 个文件，正在处理...")
        
        # 处理每个文件
        for file_name in all_files:
            # 使用 os.path.join 构建完整路径
            file_path = os.path.join(docx_dir, file_name)
            
            # 为每个文件创建单独的提取目录
            base_name = os.path.splitext(file_name)[0]
            extract_dir = os.path.join(temp_dir, base_name)
            
            try:
                os.makedirs(extract_dir, exist_ok=True)
            except Exception as e:
                print(f"创建子目录失败: {str(e)}")
                continue
            
            # 提取图片
            extracted_images, is_success = extract_images_from_docx(file_path, extract_dir)
            
            if is_success:
                # 提取成功但可能没有图片
                if extracted_images:
                    print(f"从 {file_name} 中提取了 {len(extracted_images)} 张图片")
                else:
                    print(f"从 {file_name} 中提取了 0 张图片")
            else:
                # 提取失败的提示
                print(f"■■■■■ 重要提示: 从 {file_name} 提取失败 ■■■■■")
                continue  # 提取失败，跳过后续处理
            
            # 计算每张图片的哈希值
            for img_path in extracted_images:
                img_hash = calculate_file_hash(img_path)
                if img_hash is None:
                    continue  # 跳过计算失败的图片
                    
                # 记录哈希值对应的图片信息
                hash_to_images[img_hash].append({
                    'document': file_name,
                    'image_name': os.path.basename(img_path),
                    'full_path': img_path
                })
        
        # 筛选出在多个文档中出现的图片哈希
        duplicate_hashes = {
            h: info_list for h, info_list in hash_to_images.items()
            if len(set(info['document'] for info in info_list)) > 1
        }
        
        return duplicate_hashes
        
    except Exception as e:
        print(f"处理过程中出错: {str(e)}")
        return {}

def generate_plagiarism_report(duplicates, output_file):
    """
    生成抄袭报告，统计每个文档中的抄袭图片数量，并导出到CSV文件
    """
    try:
        doc_plagiarism_count = defaultdict(set)
        
        # 遍历所有重复图片组，记录每个文档中出现的重复图片哈希
        for img_hash, info_list in duplicates.items():
            for info in info_list:
                doc_plagiarism_count[info['document']].add(img_hash)
        
        # 转换为字典，键为文档名，值为抄袭图片数量
        report_data = {doc: len(hashes) for doc, hashes in doc_plagiarism_count.items()}
        
        # 导出到CSV文件
        with open(output_file, 'w', newline='', encoding='utf-8-sig') as csvfile:
            fieldnames = ['文档名', '抄袭图片数量']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            
            writer.writeheader()
            for doc, count in sorted(report_data.items()):
                writer.writerow({'文档名': doc, '抄袭图片数量': count})
        
        print(f"\n抄袭报告已导出到: {output_file}")
        return report_data
    except Exception as e:
        print(f"生成抄袭报告失败: {str(e)}")
        return {}

def detect_image_plagiarism(docx_directory):
    """执行图片抄袭检测"""
    print("\n【第三部分】文档图片抄袭检测")
    print("=" * 50)
    
    if not os.path.isdir(docx_directory):
        print(f"错误: 目录 {docx_directory} 不存在")
        return False
    
    # 查找重复图片
    duplicates = find_duplicate_images(docx_directory)
    
    # 输出结果
    if duplicates:
        print("\n发现以下重复图片:")
        print("=" * 80)
        
        for i, (img_hash, info_list) in enumerate(duplicates.items(), 1):
            print(f"重复组 {i}:")
            print(f"  哈希值: {img_hash}")
            print(f"  出现在 {len(set(info['document'] for info in info_list))} 个文档中:")
            
            for info in info_list:
                print(f"    - 文档: {info['document']}, 图片: {info['image_name']}")
            
            print("-" * 80)
    else:
        print("未发现不同文档间的重复图片")
    
    # 生成抄袭报告并导出到CSV
    output_csv = 'plagiarism_image_detection.csv'
    generate_plagiarism_report(duplicates, output_csv)
    return True

# ========================
# 主程序入口
# ========================

def main():
    """主函数"""
    print("=" * 60)
    print("综合反抄袭检测工具")
    print("按顺序执行：文档元数据分析 -> 文字抄袭检测 -> 图片抄袭检测")
    print("=" * 60)
    
    # 提示用户输入包含Word文档的目录
    docx_directory = input("请输入文件夹路径: ").strip()
    
    # 按顺序执行三个检测步骤
    # 1. 文档元数据分析
    analyze_document_metadata(docx_directory)
    
    # 2. 文字抄袭检测
    detect_text_plagiarism(docx_directory)
    
    # 3. 图片抄袭检测
    detect_image_plagiarism(docx_directory)
    
    print("\n所有检测已完成！")

if __name__ == "__main__":
    main()