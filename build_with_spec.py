# build_with_spec.py
import os
import sys
import shutil
from PyInstaller import __main__ as pyi_main


def clean_build_dirs():
    """清理之前的构建文件"""
    dirs_to_clean = ['build', 'dist']
    for dir_name in dirs_to_clean:
        if os.path.exists(dir_name):
            print(f"清理目录: {dir_name}")
            shutil.rmtree(dir_name)


def check_required_files():
    """检查必要的文件是否存在"""
    required_files = [
        'main.py',
        'reference_processor.py',
        'app_icon.ico',
        'app_icon.png',
        'github_icon.png',
        '文献格式化工具.spec'
    ]

    missing_files = []
    for file in required_files:
        if not os.path.exists(file):
            missing_files.append(file)

    if missing_files:
        print("错误：以下文件缺失：")
        for file in missing_files:
            print(f"  - {file}")
        return False
    return True


def build_with_spec():
    """使用 spec 文件构建"""
    print("开始构建 文献引用格式化工具...")

    # 清理之前的构建
    clean_build_dirs()

    # 检查必要文件
    if not check_required_files():
        print("构建失败：缺少必要文件")
        return False

    try:
        # 使用 spec 文件构建
        print("使用 spec 文件构建...")
        pyi_main.run([
            '文献格式化工具.spec',
            '--clean'
        ])

        # 检查构建结果
        exe_path = os.path.join('dist', '文献引用格式化工具.exe')
        if os.path.exists(exe_path):
            print(f"\n✅ 构建成功！")
            print(f"可执行文件位置: {exe_path}")
            print(f"文件大小: {os.path.getsize(exe_path) / (1024 * 1024):.2f} MB")
            return True
        else:
            print("\n❌ 构建失败：未生成可执行文件")
            return False

    except Exception as e:
        print(f"\n❌ 构建过程中出现错误: {e}")
        return False


if __name__ == '__main__':
    if build_with_spec():
        print("\n🎉 程序构建完成！")
    else:
        print("\n💥 构建失败，请检查错误信息")
        sys.exit(1)