# 多格式文件 URL 和 IP 提取器
能够从多种常见文件格式（包括 TXT、CSV、Excel 和 Word）中提取 URL 和 IP 地址，并进行简单处理，输出到结果文件中。方便渗透测试时从客户提供的杂乱的文件中提取URL和IP，并去除重复数据。
## 主要功能包括：

多文件格式支持：可以处理 TXT、CSV、Excel (.xlsx/.xls) 和 Word (.docx) 文件。
数据去重：自动去除重复的 IP 和 URL，输出无冗余的结果。
结果保存：提取结果保存为 Excel 文件，便于后续分析和使用。

适用于安全分析、日志审计、数据处理等场景，是网络安全、开发人员等实用辅助工具。

## 安装说明
安装依赖环境
运行以下命令安装项目所需依赖：
```bash
pip install -r requirements.txt
```
## 使用说明
1、基本命令
```bash
python extract_ip_url.py -r 123.xlsx
```
示例：
```bash
python extract_ip_url.py -r 123.docx -o output.xlsx
```
2、参数说明：
```bash
-r 或 --read：必选参数，指定需要处理的文件路径。
-o 或 --output：可选参数，指定保存结果的文件路径，默认保存为与输入文件同名的 _result.xlsx。
```
3、运行结果
提取结果将保存到指定路径的 Excel 文件中，表格中两列分别为 IP 和 URL。
![image](https://github.com/user-attachments/assets/194c4587-da51-475c-9ca6-2d25e910b513)

