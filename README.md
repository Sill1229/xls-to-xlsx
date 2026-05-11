# xls_to_xlsx 批量转换工具

批量将 `.xls` 文件转换为 `.xlsx` 格式。  
支持真正的 Excel 二进制文件，也支持音频测试软件导出的伪装 `.xls`（实际为 TSV/CSV 文本）。

---

## 唯一前置条件：安装 Python 3

### Windows

到 [python.org/downloads](https://www.python.org/downloads/) 下载 Python 3.11+ 安装包。

> ⚠️ 安装时**务必勾选 “Add python.exe to PATH”**，否则命令行无法调用。

验证（任选一条）：

```cmd
py --version
python --version
```

显示 3.9+ 即可运行。

### macOS

打开终端执行：

```bash
brew install python@3.11
```

> 如果 `brew` 也没有，先装 Homebrew：
>
> ```bash
> /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
> ```

验证：

```bash
python3 --version
```

显示 3.9+ 即可运行。

---

**其他依赖（xlrd、openpyxl）脚本会自动安装，无需手动操作。**

---

## 使用方法

### Windows：双击 / 拖拽 `run.bat`（推荐）

- **双击 `run.bat`** —— 处理 `run.bat` 所在目录下的所有 `.xls`
- **把目标文件夹拖到 `run.bat` 图标上** —— 处理拖入目录下的所有 `.xls`
- 命令行用法：

  ```cmd
  py xls_to_xlsx.py "C:\path\to\your\folder"
  ```

### macOS / Linux：命令行

#### 方式一：指定目录

```bash
python3 xls_to_xlsx.py "/path/to/your/folder"
```

#### 方式二：把脚本放到目标目录下直接运行

```bash
cd /path/to/your/folder
python3 xls_to_xlsx.py
```

不带参数时，脚本自动处理**自身所在目录**下的所有 `.xls` 文件。

### 输出位置

转换成功的 `.xlsx` 文件保存在目标目录下的 `Convert/` 子文件夹中（自动创建）。

```
your-folder/
├── 数据A.xls
├── 数据B.xls
└── Convert/          ← 自动生成
    ├── 数据A.xlsx
    └── 数据B.xlsx
```

---

## 脚本自检功能

脚本启动时会自动执行以下检查：

| 检查项 | 通过 | 不通过时的行为 |
|--------|------|----------------|
| Python 版本 ≥ 3.9 | 继续运行 | 打印升级命令并退出 |
| xlrd 已安装 | 继续运行 | 自动 `pip install` |
| openpyxl 已安装 | 继续运行 | 自动 `pip install` |
| `.xls` 文件格式 | 按类型分流处理 | 自动识别真 Excel / 伪装 TSV/CSV |

新电脑上只要有 Python 3，直接 `python3 xls_to_xlsx.py` 即可，首次运行会自动装好依赖。

---

## 已处理的兼容性问题

| 问题 | 原因 | 脚本处理方式 |
|------|------|--------------|
| `Expected BOF record` 报错 | `.xls` 实际是 TSV/CSV 文本文件 | 自动检测文件头，分流处理 |
| 日期列显示 `########` | 列宽太窄 | 自动调整所有列宽 |
| 日期列显示为数字或无效数据 | date/time 对象缺少 Excel 格式 | 日期和时间直接写为文本字符串 |
| `zsh: command not found: pip` | macOS 下命令是 `pip3` | 脚本内部用 `sys.executable -m pip` |
| Python 版本仍为 3.9 | Xcode 自带 Python 覆盖了 Homebrew | 脚本提示添加 alias 的具体命令 |
| `No module named 'xlrd'` | 切换 Python 版本后依赖丢失 | 脚本自动检测并安装缺失依赖 |
| TSV/CSV 中文乱码（`?` 或 �） | 中文 Windows 导出的文件默认 GBK 编码，原脚本硬编码 UTF-8 | 按 utf-8-sig → utf-8 → gbk → gb18030 顺序自动回退 |
| Windows 控制台显示 emoji / 中文乱码 | cmd 默认 GBK 代码页 | `run.bat` 中 `chcp 65001`，Python 端 `sys.stdout.reconfigure` 双保险 |
| `python3` 在 Windows 上不存在 | Windows Python 安装包仅注册 `python.exe` 和 `py.exe` | `run.bat` 优先用 `py -3`，否则回退 `python` |

---

## 运行示例

```
$ python3 xls_to_xlsx.py "/Users/sly/Downloads/索尼HT-AN7"
[检查] Python 版本: 3.11.15

找到 3 个 .xls 文件，开始转换...
输出目录: /Users/sly/Downloads/索尼HT-AN7/Convert

  ✅ 不同音量频响.xls → 不同音量频响.xlsx  [实际为 TSV 文本]
  ✅ 不同音量频响dBFS.xls → 不同音量频响dBFS.xlsx  [实际为 TSV 文本]
  ✅ 漏音数据.xls → 漏音数据.xlsx  [实际为 TSV 文本]

完成: 3 成功, 0 失败
文件已保存到: /Users/sly/Downloads/索尼HT-AN7/Convert
```

### 首次运行（依赖自动安装）

```
$ python3 xls_to_xlsx.py
[检查] Python 版本: 3.11.15

[检查] 缺少依赖: xlrd, openpyxl
       正在自动安装...
       ✅ 已自动安装: xlrd, openpyxl

找到 1 个 .xls 文件，开始转换...
输出目录: /Users/sly/Downloads/Convert

  ✅ 2026-3-12_第一行理想.xls → 2026-3-12_第一行理想.xlsx  [Excel 二进制]

完成: 1 成功, 0 失败
文件已保存到: /Users/sly/Downloads/Convert
```
