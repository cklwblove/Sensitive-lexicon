# 敏感词库导出 Excel 说明

脚本 `export_sensitive_lexicon_to_excel.py` 会汇总 `ThirdPartyCompatibleFormats/TrChat` 与 `Vocabulary` 下的词库，去重合并后导出为 `敏感词库汇总.xlsx`。

## 依赖

- Python 3
- [openpyxl](https://pypi.org/project/openpyxl/)

## 安装与使用

### 方式一：项目内虚拟环境（推荐）

```bash
# 进入项目目录
cd Sensitive-lexicon

# 创建虚拟环境
python3 -m venv .venv

# 安装依赖
.venv/bin/pip install openpyxl

# 运行导出
.venv/bin/python export_sensitive_lexicon_to_excel.py
```

导出完成后，当前目录下会生成 `敏感词库汇总.xlsx`。

### 方式二：已激活虚拟环境

若已执行过 `source .venv/bin/activate`：

```bash
pip install openpyxl
python export_sensitive_lexicon_to_excel.py
```

### 方式三：全局安装 openpyxl

若本机已安装 Python 且可写 site-packages：

```bash
pip install openpyxl
python export_sensitive_lexicon_to_excel.py
```

## 输出说明

| 列     | 说明 |
|--------|------|
| 序号   | 从 1 递增 |
| 敏感词 | 词条内容 |
| 来源   | 该词出现的词库，多个用顿号分隔（如：TrChat、政治类型、色情类型） |

同一词在多个词库中出现时只占一行，来源列会合并列出。
