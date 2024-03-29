# 目的
游戏中大量使用excel配置，配置OK后需要转成lua/json/pb等格式供程序使用。
本工具依赖pb， 定义好protobuf后， 根据protobuf描述生成excel，配置excel后，通过protobuf生成lua/json/pb
程序可以以来protobuf定义，方便的读取配置。

# 使用流程
1. 设计Excel格式
2. 配置excel和转换Excel
3. 发布新配置

# 使用准备
1. 安装python3， 可以下载[最新版](https://www.python.org/ftp/python/3.12.1/python-3.12.1-amd64.exe),点击安装，注意安裝的時候要 ___勾选___ "Add python.exe to PATH"。
<img src="tools/install.png" alt="install" width="500"/>
2. pip install protobuf openpyxl

# 使用

1. 在proto目录写protobuf描述文件
2. 在tools/Config.py中配置
3. 运行gen_excel.bat生成excel
4. 在excel中增加配置
5. 运行gen_all_cfg.bat生成配置

# 工具介绍

## proto格式说明
参考tools/Config.py的配置，把proto/excel/生成文件三者关系关联起来。

```
"proto/sample.proto":
{
    "excel":"excel/sample.xlsx", 
    "serverCfg":"output/server/sample/sample.json", 
    'clientCfg':"output/client/sample/sample.json"
}
```
查看sample.proto里面的配置
<img src="tools/protobuf_tag.png" width="500"/>
- 如果某些字段只想转给服务器，就用 ***//@server***
- 如果某些字段只想转给客户端，就用 ***//@client***

生成的Excel如图所示

<img src="tools/excel.png" width="500"/>
