# IP URL domain提取工具
能够从提供的office文档中提取IP、URL、DOMAIN信息。方便渗透测试。

输出：
IPv4 → ip.txt
URL  → url.txt
域名 → domain.txt
提取结果已去重、排序，可直接导入漏扫/渗透工具

适用于渗透测试、安全分析、数据处理等场景，是网络安全、开发人员等实用辅助工具。

## 安装说明
安装依赖：
```bash
pip install -r requirements.txt
```
## 使用说明
将脚本 extract_target.py 放到与目标文件同一目录：

<img width="200" height="60" alt="image" src="https://github.com/user-attachments/assets/97048d84-e9ab-41bc-8f90-7d97d471dac1" />

1、基本命令
```bash
python extract_target.py
```
运行结束后，当前目录得到：
```bash
ip.txt      # 每行一个 IPv4
url.txt     # 每行一个 URL
domain.txt  # 每行一个 domain
```

2、扩展：
打开脚本，在「正则区域」添加自己的正则，例如提取邮箱：
```bash
EMAIL_RE = re.compile(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b')
```
然后在主逻辑里像 ip_pool 一样再建一个 email_pool 即可。

3、运行结果

<img width="500" height="60" alt="image" src="https://github.com/user-attachments/assets/b9eefb0d-127a-40fe-81bd-3372da013385" />
<img width="200" height="60" alt="image" src="https://github.com/user-attachments/assets/1a001fc8-cd3e-411d-936b-82eee824b1cb" />

4、免责声明
本工具仅供合法授权的渗透测试与安全研究使用。
一键运行，结果立得，祝各位测试顺利！


