# Excel Comparer Application

## Overview
The Excel Comparer Application is a Python-based GUI tool built using Tkinter, pandas, and SQLite. It enables users to compare two Excel files (`.xls` or `.xlsx`) and displays common rows, unique rows from each file, and the original data in a tabbed interface. The application supports Persian text rendering and right-to-left (RTL) layout for both the UI and output Excel files. Results can be saved to an SQLite database and exported to an Excel file with optional RTL formatting.

## Features
- **File Comparison**: Compares two Excel files to identify common rows and unique rows for each file.
- **Tabbed Interface**: Displays input files and comparison results in separate tabs (File 1, File 2, Common Rows, Unique Rows File 1, Unique Rows File 2).
- **Persian Language Support**: Fully supports Persian text with RTL layout using Arial font.
- **SQLite Storage**: Saves input and comparison results to an SQLite database for persistence.
- **Excel Output**: Exports results to an Excel file with optional RTL formatting for Persian text.
- **User-Friendly UI**: Includes input validation, error handling, and status updates.

## Requirements
To run the application, you need the following Python libraries:
- `tkinter` (usually included with Python)
- `pandas`
- `openpyxl` (for Excel file output)
- `sqlite3` (included with Python)

Install the dependencies using pip:
```bash
pip install pandas openpyxl
```

Ensure the Arial font is installed on your system for proper Persian text rendering.

## Usage
1. Run the script (`h.py`) using Python 3.x.
2. In the "Excel File 1" and "Excel File 2" tabs, click the respective buttons to select and load Excel files (`.xls` or `.xlsx`).
3. The application automatically compares the files and displays:
   - Original data in "Excel File 1" and "Excel File 2" tabs.
   - Common rows in the "Common Rows" tab.
   - Unique rows in the "Unique Rows (File 1)" and "Unique Rows (File 2)" tabs.
4. Check the "Set RTL Output" checkbox to enable right-to-left formatting for the output Excel file (enabled by default).
5. Click "Create Output XLS/X" to save the results to an Excel file with separate sheets for each tab's data.
6. Comparison results are automatically saved to an SQLite database (`مقایسه_اکسل.db`).

## File Structure
- `h.py`: The main Python script containing the Excel Comparer Application code.

## Notes
- The application requires the Arial font for proper Persian text rendering.
- Ensure both input Excel files have compatible data structures for accurate comparison.
- The SQLite database (`مقایسه_اکسل.db`) stores input and comparison results for persistence.
- The output Excel file includes sheets for both input files and comparison results, with optional RTL formatting.
- Large datasets are limited to 1000 rows in the UI display for performance, but all data is saved to the database and output file.

## License
This project is licensed under the MIT License.

---

# برنامه مقایسه فایل‌های اکسل

## بررسی اجمالی
برنامه مقایسه فایل‌های اکسل یک ابزار گرافیکی مبتنی بر پایتون است که با استفاده از کتابخانه‌های Tkinter، pandas و SQLite توسعه یافته است. این برنامه به کاربران امکان می‌دهد دو فایل اکسل (`.xls` یا `.xlsx`) را مقایسه کرده و ردیف‌های مشترک، ردیف‌های منحصر به فرد هر فایل و داده‌های اصلی را در یک رابط کاربری تب‌دار نمایش دهد. این برنامه از نمایش متن پارسی و چیدمان راست به چپ (RTL) برای رابط کاربری و فایل‌های خروجی اکسل پشتیبانی می‌کند. نتایج می‌توانند در یک پایگاه داده SQLite ذخیره شده و به یک فایل اکسل با قالب‌بندی اختیاری RTL صادر شوند.

## ویژگی‌ها
- **مقایسه فایل‌ها**: مقایسه دو فایل اکسل برای شناسایی ردیف‌های مشترک و ردیف‌های منحصر به فرد هر فایل.
- **رابط تب‌دار**: نمایش فایل‌های ورودی و نتایج مقایسه در تب‌های جداگانه (فایل اکسل ۱، فایل اکسل ۲، ردیف‌های مشترک، ردیف‌های منحصر به فرد فایل ۱، ردیف‌های منحصر به فرد فایل ۲).
- **پشتیبانی از زبان پارسی**: پشتیبانی کامل از متن پارسی با چیدمان راست به چپ با استفاده از فونت Arial.
- **ذخیره‌سازی در SQLite**: ذخیره ورودی‌ها و نتایج مقایسه در یک پایگاه داده SQLite برای ماندگاری.
- **خروجی اکسل**: صادر کردن نتایج به یک فایل اکسل با قالب‌بندی اختیاری RTL برای متن پارسی.
- **رابط کاربری ساده**: شامل اعتبارسنجی ورودی، مدیریت خطاها و به‌روزرسانی وضعیت.

## پیش‌نیازها
برای اجرای برنامه، به کتابخانه‌های پایتون زیر نیاز دارید:
- `tkinter` (معمولاً همراه با پایتون ارائه می‌شود)
- `pandas`
- `openpyxl` (برای خروجی فایل اکسل)
- `sqlite3` (همراه با پایتون ارائه می‌شود)

نصب وابستگی‌ها با استفاده از pip:
```bash
pip install pandas openpyxl
```

اطمینان حاصل کنید که فونت Arial روی سیستم شما نصب شده است تا متن پارسی به درستی نمایش داده شود.

## نحوه استفاده
1. اسکریپت (`h.py`) را با استفاده از پایتون 3.x اجرا کنید.
2. در تب‌های "فایل اکسل ۱" و "فایل اکسل ۲"، روی دکمه‌های مربوطه کلیک کنید تا فایل‌های اکسل (`.xls` یا `.xlsx`) را انتخاب و بارگذاری کنید.
3. برنامه به‌طور خودکار فایل‌ها را مقایسه کرده و نمایش می‌دهد:
   - داده‌های اصلی در تب‌های "فایل اکسل ۱" و "فایل اکسل ۲".
   - ردیف‌های مشترک در تب "ردیف‌های مشترک".
   - ردیف‌های منحصر به فرد در تب‌های "ردیف‌های منحصر به فرد (فایل ۱)" و "ردیف‌های منحصر به فرد (فایل ۲)".
4. گزینه "تنظیم خروجی راست به چپ (RTL)" را برای فعال کردن قالب‌بندی راست به چپ در فایل خروجی اکسل بررسی کنید (به‌طور پیش‌فرض فعال است).
5. روی دکمه "ایجاد فایل خروجی XLS/X" کلیک کنید تا نتایج در یک فایل اکسل با برگه‌های جداگانه برای داده‌های هر تب ذخیره شوند.
6. نتایج مقایسه به‌طور خودکار در یک پایگاه داده SQLite (`مقایسه_اکسل.db`) ذخیره می‌شوند.

## ساختار فایل
- `h.py`: اسکریپت اصلی پایتون حاوی کد برنامه مقایسه فایل‌های اکسل.

## نکات
- این برنامه برای نمایش صحیح متن پارسی به فونت Arial نیاز دارد.
- اطمینان حاصل کنید که هر دو فایل اکسل ورودی دارای ساختار داده‌ای سازگار برای مقایسه دقیق هستند.
- پایگاه داده SQLite (`مقایسه_اکسل.db`) ورودی‌ها و نتایج مقایسه را برای ماندگاری ذخیره می‌کند.
- فایل خروجی اکسل شامل برگه‌هایی برای هر دو فایل ورودی و نتایج مقایسه است، با قالب‌بندی اختیاری RTL.
- مجموعه‌های داده بزرگ در نمایش رابط کاربری به 1000 ردیف محدود شده‌اند تا عملکرد بهینه شود، اما تمام داده‌ها در پایگاه داده و فایل خروجی ذخیره می‌شوند.

## مجوز
این پروژه تحت مجوز MIT منتشر شده است.

---

# Excel 比较器应用程序

## 概述
Excel 比较器应用程序是一个基于 Python 的图形用户界面工具，使用 Tkinter、pandas 和 SQLite 开发。它允许用户比较两个 Excel 文件（`.xls` 或 `.xlsx`），并在选项卡界面中显示公共行、每个文件的唯一行以及原始数据。该应用程序支持波斯语文本渲染和右到左（RTL）布局，适用于用户界面和输出的 Excel 文件。结果可以保存到 SQLite 数据库，并导出到 Excel 文件，支持可选的 RTL 格式。

## 功能
- **文件比较**：比较两个 Excel 文件以识别公共行和每个文件的唯一行。
- **选项卡界面**：在单独的选项卡中显示输入文件和比较结果（文件 1、文件 2、公共行、文件 1 唯一行、文件 2 唯一行）。
- **波斯语支持**：完全支持波斯语文本，包含右到左布局，使用 Arial 字体。
- **SQLite 存储**：将输入和比较结果保存到 SQLite 数据库以实现持久化。
- **Excel 输出**：将结果导出到 Excel 文件，支持波斯语文本的可选 RTL 格式。
- **用户友好界面**：包含输入验证、错误处理和状态更新。

## 要求
运行该应用程序需要以下 Python 库：
- `tkinter`（通常随 Python 一起提供）
- `pandas`
- `openpyxl`（用于 Excel 文件输出）
- `sqlite3`（随 Python 提供）

使用 pip 安装依赖项：
```bash
pip install pandas openpyxl
```

确保系统中安装了 Arial 字体，以正确渲染波斯语文本。

## 使用方法
1. 使用 Python 3.x 运行脚本 (`h.py`)。
2. 在“Excel 文件 1”和“Excel 文件 2”选项卡中，点击相应按钮选择并加载 Excel 文件（`.xls` 或 `.xlsx`）。
3. 应用程序会自动比较文件并显示：
   - 原始数据在“Excel 文件 1”和“Excel 文件 2”选项卡中。
   - 公共行在“公共行”选项卡中。
   - 唯一行在“唯一行（文件 1）”和“唯一行（文件 2）”选项卡中。
4. 勾选“设置右到左（RTL）输出”复选框以启用输出 Excel 文件的右到左格式（默认启用）。
5. 点击“创建输出 XLS/X”按钮，将结果保存到 Excel 文件，每个选项卡的数据保存为单独的工作表。
6. 比较结果会自动保存到 SQLite 数据库 (`مقایسه_اکسل.db`)。

## 文件结构
- `h.py`：包含 Excel 比较器应用程序代码的主 Python 脚本。

## 注意事项
- 该应用程序需要 Arial 字体以正确渲染波斯语文本。
- 确保两个输入 Excel 文件具有兼容的数据结构以进行准确比较。
- SQLite 数据库 (`مقایسه_اکسل.db`) 存储输入和比较结果以实现持久化。
- 输出 Excel 文件包括输入文件和比较结果的工作表，支持可选的 RTL 格式。
- 大型数据集在用户界面显示中限制为 1000 行以优化性能，但所有数据都会保存到数据库和输出文件中。

## 许可证
本项目采用 MIT 许可证发布。