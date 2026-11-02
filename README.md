# Excel 转 iPhone 通讯录

本地 Windows 工具，把 Excel / CSV 的姓名和电话转换成 UTF-8、vCard 3.0 格式的 `.vcf` 文件。支持选择文件、拖入预览、检查错误行和导出。

版本：`1.1.0`。作者：[dadongtt](https://github.com/dadongtt)。许可证：[MIT](LICENSE)。

## 下载和使用

从 [Releases](https://github.com/dadongtt/Program-for-converting-Excel-files-to-mobile-phone-contacts/releases) 下载 Windows x64 ZIP，解压后双击 `ExcelToiPhoneVCF.exe`。下载包自带 Python，无需另外安装。

1. 准备 `.xlsx` 或 `.csv`，第一列的表头为“姓名”，第二列为“电话”。
2. 点击“选择...”或把文件拖入窗口，核对预览中的姓名、电话、状态。
3. 选择输出位置，点击“生成 VCF”。存在错误行时，程序会先询问是否只导出其余联系人。
4. 把 `.vcf` 传到 iPhone，用通讯录支持的方式导入。重复导入可能产生重复联系人。

目标使用场景是 iPhone 6s；已验证文件格式和 Windows 转换流程，尚未完成各 iOS 版本、各传输方式的真机兼容测试。

## 表格格式

| 姓名 | 电话 |
| --- | --- |
| 示例联系人001 | +12025550101 |
| 示例联系人002 | +12025550102 |

- Excel 只读取第一张工作表的前两列，第 1 行必须是表头，第 2 行开始是数据。
- 电话建议预先保存为文本，保留国际区号 `+` 和前导零。已经被 Excel 丢掉的数字无法自动找回。
- 普通数字单元格会提示改成文本；明确的连续补零格式（例如 `0000000000`）可恢复位数。
- 文字、邮箱、缺姓名、缺电话等错误行不会导出。电话检查仅检查字符和长度，不能验证号码是否存在或可拨通。
- 一行对应一个独立联系人；同名、重复号码不会自动合并或去重，姓名编号原样保留。
- 支持号码中的空格、短横线、圆括号和点，导出时会移除这些分隔符；不支持分机或一个单元格内放多个号码。
- CSV 支持 UTF-8（含 BOM）和 GB18030，采用逗号分隔。
- 提供 [示例 CSV](examples/contacts.csv)；原始测试工作簿和个人通讯录不随源码发布。

程序在本机处理输入文件，没有上传联系人或联网请求的代码。源码中的示例数据仅用于演示，请勿拨打。

## 从源码运行

建议 Windows 10/11 x64、Python 3.12（本项目在 Python 3.12.4 验证）。执行以下命令前，进入项目目录：

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe src\excel_to_iphone_vcf.py
```

也可以用 `py -3.12` 或 Python 安装路径代替 `python`。

命令行转换建议运行源码，以便看到提示：

```powershell
.\.venv\Scripts\python.exe src\excel_to_iphone_vcf.py examples\contacts.csv -o contacts.vcf
```

命令行会跳过有错误状态的行；需要逐行检查时请使用图形界面。输出路径已存在时会覆盖原文件，建议每次使用新的文件名。

## 测试和打包

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
.\.venv\Scripts\python.exe build.py
.\.venv\Scripts\python.exe tools\package_release.py
```

也可以双击 `build_exe.bat`，它优先使用项目虚拟环境，随后尝试 `py -3` 和 `python`。自定义解释器可设置 `PYTHON_EXE`。构建成功后，成品在 `release/Excel转通讯录.exe`；项目根目录的 `启动程序.bat` 是便捷入口。

构建使用 `--collect-all tkinterdnd2` 收集拖放组件。EXE 未做受信任代码签名，Windows 可能显示未知发布者；可选择自行从源码构建。正式签名可通过 `SIGN_CERT_SHA1` 指定自己的证书，并安装 Windows SDK 的 `signtool.exe`。

## 依赖与反馈

主要依赖为 Python/Tk、openpyxl、tkinterdnd2；PyInstaller 用于打包。第三方组件保持各自许可证，下载包包含其许可证文本和版本清单。

问题反馈：[Issues](https://github.com/dadongtt/Program-for-converting-Excel-files-to-mobile-phone-contacts/issues)。请说明版本、Windows 版本和复现步骤，使用虚构姓名和电话作为样本。
