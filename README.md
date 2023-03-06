# 20230303_TransferData

## exe文件的使用方法

### 文件位置：main.exe

1. 双击打开main.exe，页面提示你输入索引文件的路径：
```shell
请输入索引文件的路径，按回车键确认(默认路径为：<main.exe所在的文件夹>\工作目录\index.xls):
```
如果你什么都不填写，那么就会默认去找(<main.exe所在的文件夹>\工作目录\index.xls)

2. 请保证所有待读取的数据文件，和索引文件在同一个文件夹下。
3. 回车后，查看执行结果：


## python源代码的时候方法
1. 找到`main.py`文件，然后打开命令行
```shell
cd <main.py所在的文件夹>
pip install openpyxl xlrd # 这个命令只需要在代码第一次运行的时候执行。
python main.py
```