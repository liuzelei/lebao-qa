============================================================
  乐宝质检 - 标签批量打印工具 v2.0 使用说明
============================================================

📋 适用场景
-----------
  从金蝶云星空导出「微生物检验报告」和「成品留样保存记录」两份 Excel，
  按产品编码+料体批号自动匹配，生成 CSV 文件供 BarTender 批量打印标签。

📦 文件说明
-----------
  label_printer.py    - 主程序源码（v2.0 重构版）
  build_exe.bat       - ★ 一键打包成 exe（双击运行）
  config.ini          - 配置文件（修改 BarTender 路径、列名等）
  requirements.txt    - Python 依赖包列表
  output/             - 生成的 CSV 文件存放目录
  README.txt          - 本说明文件

🆕 v2.0 新特性
--------------
  ✦ 多线程操作 — 大文件处理时 UI 不冻结
  ✦ 进度条指示 — 操作进行中有动画提示
  ✦ 高 DPI 支持 — Windows 缩放适配，不再模糊
  ✦ 动态预览列 — 自动适配 Excel 列名，不再硬编码
  ✦ 设置对话框 — 在 GUI 中直接修改配置，无需手动编辑 config.ini
  ✦ CSV 时间戳命名 — 不再覆盖历史数据（如 print_data_20260629_120000.csv）
  ✦ 未匹配记录过滤 — 勾选「仅打印匹配成功记录」自动排除
  ✦ 打印选中记录 — 在预览表中选中行后可单独打印
  ✦ 日志系统 — 所有操作记录到 label_printer.log
  ✦ 快捷键 — Ctrl+M 匹配 / Ctrl+G 生成 / Ctrl+P 打印 / Ctrl+T 测试
  ✦ 统一 ttk/clam 主题 — 视觉一致的现代风格 + 彩色按钮
  ✦ BarTender 命令统一格式 — 引号处理一致，减少打印错误
  ✦ 优雅依赖缺失提示 — 缺 pandas 时弹窗提示而非闪退
  ✦ 单次 Excel 读取 — 验证+匹配一次完成，不再重复读文件

🚀 推荐方式：一键打包成独立 EXE（无需安装 Python）
----------------------------------------------------

  前提：Windows 电脑上装好 Python 3.8+
        （仅打包时需要，生成后的 exe 可在任何电脑直接运行）

  步骤：
    ① 将整个 label_printer/ 文件夹拷贝到 Windows 电脑
    ② 双击运行 build_exe.bat
    ③ 等待 2-5 分钟，自动完成打包
    ④ 打包完成后，dist\乐宝标签打印工具\ 文件夹会自动打开
    ⑤ 修改 dist\乐宝标签打印工具\config.ini 中的路径配置
    ⑥ 双击 乐宝标签打印工具.exe 即可使用

  生成的 dist\乐宝标签打印工具\ 文件夹可以拷贝到任意 Windows 电脑直接运行，
  不需要安装 Python。如需分发，建议压缩为 zip。

  注: v2.0 使用 --onedir 打包模式，启动速度比旧版 --onefile 快 2-3 倍。

⚙️ 备选方式：直接用 Python 运行（开发调试用）
----------------------------------------------

  1. 安装 Python 3.8+
     https://www.python.org/downloads/
     ⚠️ 安装时勾选 "Add Python to PATH"

  2. 安装依赖
     pip install pandas openpyxl

  3. 修改 config.ini（路径配置见下方）

  4. 双击 label_printer.py 运行

📝 配置文件说明
---------------
  用记事本打开 config.ini，或在程序中点击「文件 → 设置」(Ctrl+S) 修改：

  [BarTender]
  bartender_path = C:\Program Files\Seagull\BarTender 9.4\bartend.exe
  template_path = D:\标签模板\微生物检测报告标签.btw
  printer_name = TSC TTP-342E Pro
  copies = 1

  [Excel]
  report_key_product = 产品编码          ← 微生物报告中的产品编码列名
  report_key_batch = 样品批号            ← 微生物报告中的批号列名
  sample_key_product = 产品代码          ← 留样记录中的产品编码列名
  sample_key_batch = 料体批号            ← 留样记录中的批号列名

  请将 bartender_path 和 template_path 改成你电脑的实际路径。
  如果 Excel 列名不同，也一并修改（或通过设置对话框修改）。

⚙️ 操作步骤（共4步）
--------------------

  第①步：选择文件
    ① 双击 exe 启动程序
    ② 点击「浏览...」选择微生物检验报告 Excel
    ③ 点击「浏览...」选择成品留样保存记录 Excel

  第②步：检查格式并匹配（★ 关键步骤）
    ④ 点击绿色的「① 检查格式并匹配」按钮（或 Ctrl+M）
    ⑤ 软件自动检查两份Excel的格式，弹窗显示检查报告：
       - 有错误(✗) → 停止操作，修正Excel后重新选择文件
       - 有警告(⚠) → 弹窗询问是否继续（建议先检查文件）
       - 全部通过(✓) → 自动开始匹配数据
    ⑥ 匹配完成后弹窗显示结果

  第③步：生成 CSV
    ⑦ 点击蓝色的「② 生成CSV」按钮（或 Ctrl+G）
    ⑧ 可勾选「仅打印匹配成功的记录」过滤未匹配数据
    ⑨ CSV 文件保存在 exe 同目录下的 output/ 文件夹
       文件名包含时间戳，不会覆盖历史数据

  第④步：打印
    ⑩ 先点「测试打印(首条)」(Ctrl+T) 打一条，检查标签效果
    ⑪ 确认无误后，点「③ 批量打印」(Ctrl+P)
    ⑫ 如需只打印部分记录，在预览表中选中行后点「打印选中(N条)」

  💡 建议：每次都先测试打印第一条，确认标签内容正确后再批量打印。

🔍 数据格式检查说明
--------------------
  点击「检查格式并匹配」后，软件会自动检查以下内容：

  微生物检验报告 Excel:
    ✓ 文件能否正常打开
    ✓ 是否包含「产品编码」列
    ✓ 是否包含「样品批号」列
    ⚠ 关键列是否有空值
    ⚠ 「检测结论」是否为空

  成品留样保存记录 Excel:
    ✓ 文件能否正常打开
    ✓ 是否包含「产品代码」列
    ✓ 是否包含「料体批号」列
    ⚠ 关键列是否有空值

  如果检查发现问题，会弹窗显示详细报告，不会盲目匹配。

⌨️ 快捷键
---------
  Ctrl+M     检查格式并匹配
  Ctrl+G     生成CSV
  Ctrl+P     批量打印
  Ctrl+T     测试打印（首条）
  Ctrl+O     打开检验报告文件
  Ctrl+S     打开设置对话框

🔧 常见问题
-----------
  Q: build_exe.bat 打包失败？
  A: ① 确认 Python 3.8+ 已安装且加入 PATH
     ② 暂时关闭杀毒软件
     ③ 确保磁盘有 500MB 以上空间

  Q: 程序启动时弹窗提示"缺少依赖库"？
  A: 运行 pip install pandas openpyxl 安装依赖后重试。
     打包为 exe 后则不需要此步骤。

  Q: 提示缺少列名？
  A: 修改 config.ini 的 [Excel] 部分，或在程序中「文件 → 设置」修改。
     与你的 Excel 表头保持一致。

  Q: 打印后标签内容为空？
  A: 检查 config.ini 中 template_path 是否正确。

  Q: 提示找不到 BarTender？
  A: 检查 bartender_path，BarTender 9.4 通常装在：
     C:\Program Files\Seagull\BarTender 9.4\bartend.exe

  Q: CSV 中文乱码？
  A: 本工具使用 UTF-8 with BOM，BarTender 9.4 应能识别。
     如仍有问题请反馈。

  Q: 程序崩溃/报错？
  A: 查看 exe 同目录下的 label_printer.log 日志文件，
     包含详细错误信息，可发给技术支持分析。

  Q: 想切换为单文件 exe（慢启动但方便分发）？
  A: 将 build_exe.bat 中的 --onedir 改为 --onefile 即可。
