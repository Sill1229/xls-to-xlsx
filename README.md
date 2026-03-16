# xls_to_xlsx 批量转换工具

批量将 `.xls` 文件转换为 `.xlsx` 格式。  
支持真正的 Excel 二进制文件，也支持音频测试软件导出的伪装 `.xls`（实际为 TSV/CSV 文本）。

---

## 唯一前置条件：安装 Python 3

macOS 用户打开终端执行：

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

显示 3.9+ 即可运行。**其他依赖（xlrd、openpyxl）脚本会自动安装，无需手动操作。**

---

## 使用方法

### 方式一：指定目录

```bash
python3 xls_to_xlsx.py "/path/to/your/folder"
```

### 方式二：把脚本放到目标目录下直接运行

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

## 运行示例

```
$ python3 xls_to_xlsx.py "/Users/sly/Downloads/索尼HT-AN7"
[检查] Python 版本: 3.11.15
[检查] 缺少依赖: xlrd, openpyxl
       正在自动安装...
       ✅ 已自动安装: xlrd, openpyxl

找到 3 个 .xls 文件，开始转换...
输出目录: /Users/sly/Downloads/索尼HT-AN7/Convert

  ✅ 不同音量频响.xls → 不同音量频响.xlsx  [实际为 TSV 文本]
  ✅ 不同音量频响dBFS.xls → 不同音量频响dBFS.xlsx  [实际为 TSV 文本]
  ✅ 漏音数据.xls → 漏音数据.xlsx  [实际为 TSV 文本]

完成: 3 成功, 0 失败
文件已保存到: /Users/sly/Downloads/索尼HT-AN7/Convert
```
