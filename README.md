# 聊天记录合规关键词筛查工具

内网 Windows 单机运行的桌面小工具：上传聊天记录 Excel、粘贴一批关键词，自动标出命中的聊天记录。

## 一、怎么用（exe 版，开箱即用）

1. 双击 `ChatAuditTool.exe`
2. **① 选择文件**：点「浏览…」，选待检查的 `.xlsx`
3. **② 填关键词**：把关键词粘进去，用「、」分隔（也兼容换行、逗号、分号）；可从 txt 导入；关键词会自动存在本机，下次打开自动载入
4. **③ 开始筛查**：勾「区分大小写」则严格匹配大小写，默认忽略大小写
5. **④ 看结果**：双击任意一行 → 弹窗显示完整原文，命中的关键词黄色高亮
6. **⑤ 导出**：
   - **导出标记版 Excel**：复制原文件，在被扫描的工作表末尾新增「是否命中」「命中的关键词」两列（原文件不改动）
   - **仅导出命中行**：只保留命中的行，另存一份新文件
   - **命中统计**：按关键词汇总命中次数

## 二、Excel 要求

工具按**列名**自动识别，识别不到时按位置回退（第 1 个工作表 F 列、第 2 个工作表 B 列）：

| 工作表 | 识别依据 | 参与扫描的列 |
|---|---|---|
| 业务部门类 | 表头含「消息内容」 | 消息内容 |
| 贵金属类 | 表头含「聊天记录」 | 聊天记录 |
| 人员清单等 | 无上述列名 | 不扫描 |

- 第 2 张表的「行员姓名」是合并单元格，工具会自动向下填充
- 空行自动跳过
- **注意**：Excel 单元格最多存 32767 字，超长聊天会被截断。命中的行若达到该长度，界面会提示「⚠ 疑似被 Excel 截断」，被截掉的部分无法检索到（需源头导出时拆分）

## 三、性能

匹配量与字符总量相关：实测样例 11058 行 × 318 个关键词 ≈ **1.6 秒**；1 万行 × 500 词约 3–8 秒。扫描在后台线程，界面不卡，可随时「停止」。

## 四、怎么拿到 exe（GitHub Actions 打包）

本机是 macOS，PyInstaller 不能交叉编译，所以用 GitHub 的 Windows  runner 打包：

1. 把本目录推到 GitHub 仓库
2. 打开仓库 → **Actions** → `Build Windows EXE` → **Run workflow**（或推一个 `v1.0` 标签自动触发）
3. 跑完后在 Artifacts 里下载 `ChatAuditTool-windows-exe`，解压即得 `ChatAuditTool.exe`

如需本地打包（Windows 机器）：

```bat
pip install -r requirements.txt
pip install pyinstaller
pyinstaller --noconsole --onefile --clean --name ChatAuditTool main.py
```

## 五、源码结构

```
main.py      图形界面（tkinter）
scanner.py   工作表/列识别 + 关键词扫描
exporter.py  导出标记版 Excel / 命中清单
config.py    关键词本机持久化（~/.chat_audit_config.json）
```

本地运行：`pip install -r requirements.txt && python main.py`

> ⚠ macOS 本机验证时注意：系统自带的 `/usr/bin/python3`(3.9) 绑定的是 **Tk 8.5**，ttk 控件会渲染不出来（窗口只有标题、内容空白）。请用 Tk 8.6+ 的解释器，例如 python.org 官方 Python 3.11+ 或 `python3.11 -m venv .venv`。Windows 打包环境（Python 3.11/Tk 8.6）无此问题。
