import os
import sys
import subprocess

# 图标文件路径
icon_path = r"C:\Users\4_of_Diamonds\Pictures\Camera Roll\images.png"
# 输出文件名
exe_name = "WordExtractor"
# 主程序文件
main_file = "word_extractor.py"

# 检查图标文件是否存在
if not os.path.exists(icon_path):
    print(f"错误: 图标文件 {icon_path} 不存在!")
    sys.exit(1)

# 构建PyInstaller命令
cmd = [
    "pyinstaller",
    "--noconfirm",  # 覆盖输出目录
    "--windowed",   # 不显示控制台窗口
    "--onefile",    # 打包成单个可执行文件
    f"--icon={icon_path}",  # 设置图标
    f"--name={exe_name}",   # 设置输出文件名
    "--clean",      # 清理临时文件
    "--add-data=README.md;.",  # 添加README文件
    # 添加隐式导入，确保所有依赖都被包含
    "--hidden-import=docx",
    "--hidden-import=win32com",
    "--hidden-import=win32com.client",
    main_file
]

print("开始打包Word文档页面提取器...")
print(f"使用图标: {icon_path}")
print(f"输出文件: {exe_name}.exe")

try:
    # 执行PyInstaller命令
    subprocess.run(cmd, check=True)
    print("\n打包完成!")
    print(f"可执行文件位于: {os.path.join('dist', exe_name + '.exe')}")
except subprocess.CalledProcessError as e:
    print(f"\n打包过程中出错: {e}")
    sys.exit(1)
except Exception as e:
    print(f"\n发生未知错误: {e}")
    sys.exit(1)