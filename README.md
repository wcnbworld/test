# 专利表格分类软件

一个基于 Streamlit 的小工具：
- 左侧为输入窗口（上传表格A）；
- 右侧为输出窗口（显示分类结果并下载 Word 文档）。

## 运行方式（源码）

```bash
pip install -r requirements.txt
streamlit run app.py
```

## 直接打包 EXE（Windows）

### 方法1：一键脚本
在 Windows 命令行中运行：

```bat
build_windows_exe.bat
```

成功后会在 `dist/PatentClassifier.exe` 生成可执行文件。

### 方法2：手动命令

```bat
pip install -r requirements.txt
pip install pyinstaller
pyinstaller --noconfirm --clean --onefile --name PatentClassifier launcher.py
```

> 注意：EXE 需要在 Windows 环境中打包（建议在目标系统或同版本系统打包）。

## 功能说明

- 读取包含专利字段的 Excel/CSV。
- 标题(译)(简体中文)、独立权利要求、技术功效、摘要(译)(简体中文)按“或”关系处理，四者至少有一列即可运行。
- 自动把申请人按给定企业映射归一化。
- 基于标题、独立权利要求、技术功效、摘要进行技术分类：
  - 赛车模拟器
  - 飞行模拟器
  - 手柄
  - 渔线轮
  - 电动扳手
  - 其他
- 输出两个 Word：
  - 主文档：按申请人->专利类型生成表格B（外观）/表格C（发明、授权发明、实用新型）。
  - 表格D（其他分类）在行数 < 100 时输出 Word；行数 >= 100 时自动输出 Excel。
- 从表格A中提取“摘要附图”并写入表格B/C对应“专利图片”列（xlsx支持图片提取）。
- “申请信息”列仅显示内容，不显示字段名前缀；申请人显示原始申请人（非归一化名称）。
- 表格B/C/D中的图片高度统一为 2.7 厘米，并保持纵横比。
