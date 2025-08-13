#!/usr/bin/env python3
import csv
import os
from pathlib import Path


def extract_projects_from_csv(csv_file_path):
    """从CSV文件提取唯一的项目名称"""
    projects = set()
    
    with open(csv_file_path, 'r', encoding='utf-8') as file:
        reader = csv.DictReader(file)
        for row in reader:
            project = row.get('所属项目', '').strip()
            if project:
                projects.add(project)
    
    return sorted(list(projects))


def save_projects_to_file(projects, output_file_path):
    """保存项目名称到文件"""
    with open(output_file_path, 'w', encoding='utf-8') as file:
        for project in projects:
            file.write(project + '\n')


def create_project_directories(base_path, projects):
    """在指定路径下创建项目目录"""
    new_dir_path = Path(base_path) / 'new_dir'
    new_dir_path.mkdir(exist_ok=True)
    
    for project in projects:
        project_dir = new_dir_path / project
        project_dir.mkdir(exist_ok=True)


def manage_project_structure(csv_file_path='document_tags.csv', base_path='.'):
    """主要管理方法：提取项目、保存文件、创建目录"""
    # 提取项目名称
    projects = extract_projects_from_csv(csv_file_path)
    
    # 保存到mkdir_csv文件
    mkdir_csv_path = Path(base_path) / 'mkdir_csv'
    save_projects_to_file(projects, mkdir_csv_path)
    
    # 创建项目目录
    create_project_directories(base_path, projects)


if __name__ == "__main__":
    manage_project_structure()