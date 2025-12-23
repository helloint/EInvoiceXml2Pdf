# 全电发票 XML 转 PDF
![1-标题-2.png](./docs/images/EInvoicePDF.png)

## 使用方法
本项目为C#工程，可以用VS或VSCode打开。

如果用VSCode打开，需要安装.net SDK。如果用的是Mac，还会提示你安装Mac的Runtime，按照提示来就行了。

把需要转换的电子发票XML文件放在`/assets/`目录下，转换成的PDF会输出到`/output/`目录下。

然后：
```shell
# 先编译再运行
dotnet build
dotnet run --project EInvoiceXml2Pdf
```

## 备注
本工程为Fork https://github.com/BitBrewing/EInvoiceXml2Pdf/fork 后用AI做了增强，支持批量转换，输出到根目录，方便使用。

## 引用
- [iTextSharp.LGPLv2.Core](https://github.com/VahidN/iTextSharp.LGPLv2.Core)

## 参考资料
- [财政部会计司关于公布电子凭证会计数据标准（试行版）的通知](https://kjs.mof.gov.cn/zt/kuaijixinxihuajianshe/dzpzkjsjbzshsd/sjbz/202305/t20230517_3885004.htm)