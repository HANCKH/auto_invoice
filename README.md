# auto_invoice

这是一个用于发票整理的仓库，当前主要包含发票处理脚本和模板文件。

## 文件说明

- `invoice_to_excel.py`
  - 发票批量处理脚本。
  - 可读取指定目录下的 PDF 发票，提取公司、明细、金额、税额、发票号等信息。
  - 支持按发票号去重。
  - 支持按公司归档输出：每个公司一个文件夹、一个汇总 Excel。
  - 支持将同公司的多张发票按“明细行”写入同一个 Excel。
  - 支持生成总表 `A物品清单.xlsx`。
  - 支持可选重命名/复制 PDF 为 `<公司名>_<金额>.pdf`。

- `模板文件.xlsx`
  - 导入模板文件。
  - 脚本会根据模板 `导入` 工作表的表头动态匹配写入，兼容新旧列顺序差异。

- `.gitignore`
  - 控制仓库仅跟踪需要的核心文件。

## 脚本使用方式

```bash
python invoice_to_excel.py \
  --input-dir <发票目录> \
  --template 模板文件.xlsx \
  --output-dir <发票目录>/整理结果 \
  --rename
```

常用参数：

- `--input-dir`：输入发票目录
- `--template`：模板文件路径
- `--output-dir`：输出目录
- `--rename`：输出重命名后的 PDF
- `--recursive`：递归扫描子目录中的 PDF

## 输出结果说明

在 `<输出目录>` 下会生成：

1. 按公司名称分组的子文件夹
2. 每个公司一个汇总 Excel（包含该公司全部发票明细）
3. 对应的 PDF 文件
4. 总表 `A物品清单.xlsx`

## 依赖

```bash
pip install pdfplumber openpyxl pypinyin
```
