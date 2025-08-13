#!/usr/bin/env python3
import csv
import os
import shutil
from pathlib import Path


def read_csv_mapping(csv_file_path):
    """从CSV文件读取文档标题到项目的映射关系"""
    mapping = {}
    
    with open(csv_file_path, 'r', encoding='utf-8') as file:
        reader = csv.DictReader(file)
        for row in reader:
            title = row.get('文档标题', '').strip()
            project = row.get('所属项目', '').strip()
            if title and project:
                mapping[title] = project
    
    return mapping


def find_document_files(base_path):
    """查找项目中所有支持的文档文件"""
    supported_extensions = ['.txt', '.docx']
    document_files = []
    
    base_path = Path(base_path)
    
    # 搜索根目录和documents子目录
    search_paths = [base_path, base_path / 'documents']
    
    for search_path in search_paths:
        if search_path.exists():
            for ext in supported_extensions:
                document_files.extend(search_path.rglob(f'*{ext}'))
    
    return document_files


def move_documents_to_projects(base_path, mapping):
    """根据映射关系移动文档到对应项目目录"""
    base_path = Path(base_path)
    new_dir_path = base_path / 'new_dir'
    
    # 查找所有文档文件
    document_files = find_document_files(base_path)
    
    for file_path in document_files:
        # 获取文件名（不含扩展名）
        file_title = file_path.stem
        
        # 查找映射关系
        if file_title in mapping:
            project = mapping[file_title]
            target_dir = new_dir_path / project
            target_file = target_dir / file_path.name
            
            # 移动文件
            if target_dir.exists():
                shutil.copy2(str(file_path), str(target_file))


def move_documents_by_csv(csv_file_path='document_tags.csv', base_path='.'):
    """主要封装方法：根据CSV文件移动文档到对应项目目录"""
    # 读取CSV映射关系
    mapping = read_csv_mapping(csv_file_path)
    
    # 移动文档
    move_documents_to_projects(base_path, mapping)


if __name__ == "__main__":
    move_documents_by_csv()