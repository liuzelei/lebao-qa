# 乐宝质检 - 原料留样标签批量打印工具

从金蝶云星空导出物料数据 Excel，逐条打印标签（通过 BarTender CLI）。

## 功能

- **一个 Excel → 逐条打印**：选择一个数据文件 + 一个 BTW 模板即可
- **数据质量检查**：打印前自动检查必检列空值，提示具体行号（如"列「编码」第 3, 7 行为空"）
- **空值行橙色高亮**：预览表中空值行自动标橙色背景
- **逐条打印带进度**：进度条显示"正在打印 3/10 (30%)"
- **模板一次出两张**：BTW 模板为 2-up 布局，每条记录打印 2 张标签（copies=1）
- **模板路径 GUI 可选**：界面上有浏览按钮选 .btw 文件，不需要手动改 config.ini
- **CSV 时间戳命名**：每次生成带时间戳的 CSV，不覆盖历史数据

## 快速开始

### 方式 1：下载 Release（推荐，无需安装 Python）

1. 从 [GitHub Releases](https://github.com/liuzelei/lebao-qa/releases) 下载对应架构的 zip：
   - **LebaoLabelPrinter-x64.zip**：适用于绝大多数 Windows 电脑（Intel/AMD 64位），**支持 Windows 7 SP1 及以上**
   - **LebaoLabelPrinter-arm64.zip**：适用于 ARM64 Windows 设备（需 Windows 11）
2. 解压到任意目录
3. 编辑 `config.ini` 配置打印机名称和必检列名
4. 双击 `LebaoLabelPrinter.exe` 运行

### 方式 2：本地打包成 EXE

前提：Windows 电脑上安装了 Python 3.8-3.13

1. 将整个 `label_printer/` 文件夹拷贝到 Windows 电脑
2. 双击运行 `build_exe.bat`（自动检测 Python 版本，Win7 优先选 3.8）
3. 等待 2-5 分钟，打包完成后自动打开 `dist\LebaoLabelPrinter\` 目录
4. 修改 `config.ini` 中的路径配置
5. 双击 `LebaoLabelPrinter.exe` 即可使用

> 打包出的 `dist\LebaoLabelPrinter\` 文件夹可以拷贝到任意 Windows 电脑直接运行，不需要安装 Python。

### 方式 3：直接用 Python 运行（开发调试用）

```bash
pip install pandas openpyxl
python label_printer.py
```

## 配置文件说明

用记事本打开 `config.ini` 修改：

```ini
[BarTender]
bartender_path = C:\Program Files\Seagull\BarTender 9.4\bartend.exe
template_path =                  ; 留空则在 GUI 中选择
printer_name = TSC TTP-342E Pro
copies = 1                       ; BTW 2-up 模板 copies=1 即每条出2张

[Excel]
required_columns = 编码,名称,规格型号,基本单位   ; 必检列名（逗号分隔）
```

> `template_path` 可以留空，在 GUI 界面通过浏览按钮选择。选择后会自动保存到 config.ini。

## 操作步骤

1. **选择文件**：点击"浏览..."选择 Excel 数据文件
2. **选择模板**：点击"浏览..."选择 BarTender 模板 (.btw)
3. **检查数据**：点击"① 检查数据"，自动检查格式和空值，预览数据
4. **测试打印**：建议先点"测试首条"打一条确认标签效果
5. **批量打印**：确认无误后点"② 打印标签"，逐条打印带进度

## 文件说明

| 文件 | 说明 |
|------|------|
| `label_printer.py` | 主程序源码 |
| `build.py` | Python 构建脚本（自动检测版本、打包 vcruntime DLL） |
| `build_exe.bat` | 一键打包（双击运行，自动选 Python 版本，Win7 优先选 3.8） |
| `config.ini` | 配置文件 |
| `requirements.txt` | Python 依赖：pandas, openpyxl |
| `output/` | 生成的 CSV 文件存放目录 |

## 构建架构说明

| 架构 | Python 版本 | pandas 版本 | Windows 最低版本 |
|------|-------------|-------------|-----------------|
| x64 | 3.8 | <2.0 (1.5.x) | Windows 7 SP1 |
| arm64 | 3.13 | latest | Windows 11 |

> Python 3.9+ 依赖 `api-ms-win-core-path-l1-1-0.dll`（Win8+ API），所以在 x64 构建中使用 Python 3.8 以确保 Win7 兼容性。

## 常见问题

**Q: build_exe.bat 打包失败？**
A: ① 确认 Python 3.8-3.13 已安装 ② 暂时关闭杀毒软件 ③ 确保磁盘有 500MB 以上空间

**Q: 程序启动时弹窗提示"缺少依赖库"？**
A: 运行 `pip install pandas openpyxl` 安装依赖。打包为 exe 后不需要此步骤。

**Q: 提示缺少列名？**
A: 修改 `config.ini` 的 `[Excel]` 部分 `required_columns`，与你的 Excel 表头保持一致。

**Q: 打印后标签内容为空？**
A: 检查 BTW 模板中的字段映射是否与 Excel/CSV 列名对应。

**Q: 提示找不到 BarTender？**
A: 检查 `config.ini` 中 `bartender_path` 是否正确。BarTender 9.4 通常装在：
`C:\Program Files\Seagull\BarTender 9.4\bartend.exe`

**Q: CSV 中文乱码？**
A: 本工具使用 UTF-8 with BOM 编码，BarTender 9.4 应能识别。如仍有问题请反馈。

**Q: Win7 上运行提示缺少 DLL？**
A: 确保 x64 版本是用 Python 3.8 构建的（GitHub Release 的 x64 包已兼容 Win7）。

**Q: 程序崩溃/报错？**
A: 查看 `label_printer.log` 日志文件，包含详细错误信息。
