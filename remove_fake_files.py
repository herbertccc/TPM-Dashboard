#!/usr/bin/env python3
"""
彻底清理假数据：删除假CSV文件，只保留真实数据源
"""
import os, shutil

BASE_DIR = '/Users/herbert/Desktop/Project_Report_v2'
DATA_DIR = os.path.join(BASE_DIR, 'data')

# 需要删除的假数据文件
FAKE_FILES = ['E1W_App.csv', 'E2W_Web.csv', 'E3W_API.csv']

def main():
    print("=== 假数据清理 ===\n")
    
    # 1. 删除假CSV文件
    print("1. 删除假数据文件:")
    for fname in FAKE_FILES:
        fpath = os.path.join(DATA_DIR, fname)
        if os.path.exists(fpath):
            os.remove(fpath)
            print(f"   ✓ 已删除: {fname}")
        else:
            print(f"   - 不存在: {fname}")
    
    # 2. 列出剩余真实文件
    print("\n2. data/ 目录剩余文件:")
    remaining = [f for f in os.listdir(DATA_DIR) if not f.startswith('.')]
    for f in remaining:
        print(f"   • {f}")
    
    # 3. 更新 generate.py 中的 auto_discover_projects 逻辑
    #    由于只剩一个真实文件，可以简化为显式列表或保持自动扫描
    print("\n3. generate.py 将自动扫描剩余文件（仅【E1W-APP项目】.xlsx）")
    
    # 4. 提示用户关于里程碑映射
    print("\n4. 注意: milestone/里程碑节点.xlsx 包含 B30/X50 两个sheet")
    print("   如果这些也属于假数据，请手动删除 milestone/ 目录")
    print("   当前保留该文件，但映射关系可能需要调整")
    
    print("\n✓ 假数据文件清理完成")
    print(f"  真实数据源: {remaining}")

if __name__ == '__main__':
    main()
