# 营业额统计工具

一个用于合并“钉钉记录”和“日流水表”的 Windows 桌面工具，可自动匹配门店数据并生成统一的 Excel 营业额统计表。

## 功能

- 合并钉钉记录与日流水数据
- 自动匹配门店并处理重复记录
- 对异常金额进行日志提醒
- 生成格式化 Excel 统计表
- Win11 风格中文桌面界面
- 支持水平拖动调整窗口宽度
- 支持图形界面和命令行两种运行方式

## 直接使用

前往仓库的 [Releases](../../releases) 下载最新版 Windows EXE，无需安装 Python。

启动程序后：

1. 选择钉钉记录 Excel 文件。
2. 选择日流水 Excel 文件。
3. 可选填写统计日期和输出位置。
4. 点击“生成并导出 Excel”。

## 从源码运行

需要 Python 3.10 或更高版本：

```powershell
pip install -r requirements.txt
python 营业额统计生成.py
```

命令行模式：

```powershell
python 营业额统计生成.py --cli --dingding 钉钉记录.xlsx --riliushui 日流水.xls --output 统计表.xlsx
```

## 打包 Windows EXE

```powershell
pip install pyinstaller
pyinstaller --noconfirm --clean --onefile --windowed `
  --name "营业额统计工具" `
  --icon "app-icon.ico" `
  --add-data "app-icon.ico;." `
  营业额统计生成.py
```

## 发布命名

GitHub Release 统一命名为：`美优乐营业额统计工具 v版本号`。

## 数据说明

- 日流水作为原始财务数据来源。
- 钉钉记录用于补充缺失数据。
- 两份数据均有非零值且不一致时，以日流水为准。
- 工具仅在本机处理文件，不会上传营业数据。
