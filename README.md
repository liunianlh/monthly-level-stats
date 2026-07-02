# 月度等级统计程序

用于读取固定格式的 Excel 统计表，按“过犹”日期的月份统计 A、B、C、D、E 等级数量。

## 使用源码运行

1. 安装 Python 3.14，并确保安装了 Tcl/Tk（Python.org 的 Windows 标准安装包默认包含）。
2. 在本目录执行：

   ```bat
   py -3.14 -m pip install -r requirements.txt
   py -3.14 app.py
   ```

3. 点击“选择 Excel”，选择 `.xls` 或 `.xlsx` 文件。
4. 点击“开始统计”，查看结果。
5. 点击“导出 Excel”，保存汇总文件。

## 打包 Windows EXE

1. 将整个文件夹复制到 Windows 10/11 64 位电脑。
2. 安装 Python 3.14，并勾选“Python Launcher”和 Tcl/Tk 相关组件。
3. 双击 `build_windows.bat`。
4. 脚本会安装依赖、运行测试并生成：

   ```text
   dist\MonthlyLevelStats.exe
   ```

脚本会强制使用 Python 3.14，并在 `.venv314` 中安装 PyInstaller 6.21.0。生成后可将 `MonthlyLevelStats.exe` 单独复制到其他 Windows 10/11 64 位电脑运行，无需安装 Python。

## Excel 模板要求

- 文件格式必须为 `.xls` 或 `.xlsx`。
- 必须包含程序所需的统计工作表。
- `A1` 必须为“等级”，明细只能填写 A、B、C、D、E。
- `C1` 必须为“过犹”。
- 过犹日期格式必须类似：`保单过犹日期：2025-06-19`。
- 每条明细行计数一次，不按姓名、身份证或保单号去重。

程序遇到不符合模板的数据会指出具体错误和行号，不会静默忽略。
