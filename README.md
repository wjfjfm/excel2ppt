# Excel2PPT script

Generating pptx(s) with an excel to provide data and a pptx as template.

## Install

Make sure you have python3 installed first:

```
python --version
```

Install dependencies:

```
pip install python-pptx
pip install openpyxl
pip install pandas
```

## How to use

1. edit the data.xlxs and template.pptx

2. `python excel2ppt.py`

Done!

## For Example

`data.xlsx` like this:

```
filename	text1	text2	text3
张三的文件	张三	1234	A部门
李四的文件	李四	12345	B部门
王五的文件	王五	123456	C部门
```

`template.pptx` like this:

![image-20220323151649617](README.assets/image-20220323151649617.png)

after `python excel2ppt.py`

```
$ python excel2ppt.py
openning data.xlsx ...
found var: filename
found var: text1
found var: text2
found var: text3
openning template.pptx ...
Done
generating 张三的文件.pptx
Done!
generating 李四的文件.pptx
Done!
generating 王五的文件.pptx
Done!
Everything Done, exiting ...
```

Generates 3 file, like this:

![image-20220323151857009](README.assets/image-20220323151857009.png)
