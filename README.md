---
title: 使用Python读取Excel中的图片并对应到记录
date: 2021-11-16 21:40:58
tags:
  - Python
  - Excel
categories:
  - Python
---

## 1、需求

有这样一个需求，将采集在 Excel 中的人员信息（包含照片）导入到 Mysql 库。需求很简单，就是读取 Excel 中的数据插入到 Mysql 中的表，问题在于照片怎么读取？

## 2、分析

对于读取文本数据，直接按行读取即可；对于图片，常用做法是，将 Excel 文件的后缀名改为 zip，然后解压文件，对应文件名目录下有一个`\xl\media`的目录，里面便是我们要照片的图片，然而，他的文件名、文件顺序都是乱的，确定不出图片对应的记录，怎么办呢？其实目录中还有这样一个 xml 文件`\xl\drawings\drawing1.xml`其中就有图片与记录的对应关系，只要解析 xml 文件就可以了。知道了这个思路，那我们来一步步的实现他。

## 3、实现

1、复制并修改指定目录下的文件类型名，将excel后缀名修改为.zip

```python
def copy_change_file_name(file_path, new_type='.zip'):
    if not isfile_exist(file_path):
        return ''
    extend = os.path.splitext(file_path)[1]  # 获取文件拓展名
    if extend != '.xlsx' and extend != '.xls':
        print("It's not a excel file! %s" % file_path)
        return False
    file_name = os.path.basename(file_path)  # 获取文件名
    new_name = str(file_name.split('.')[0]) + new_type  # 新的文件名，命名为：xxx.zip

    dir_path = os.path.dirname(file_path)  # 获取文件所在目录
    new_path = os.path.join(dir_path, new_name)  # 新的文件路径
    if os.path.exists(new_path):
        os.remove(new_path)
    shutil.copyfile(file_path, new_path)
    return new_path  # 返回新的文件路径，压缩包
```

2、解压文件

```python
def unzip_file(zipfile_path):
    if not isfile_exist(zipfile_path):
        return False
    if os.path.splitext(zipfile_path)[1] != '.zip':
        print("It's not a zip file! %s" % zipfile_path)
        return False
    file_zip = zipfile.ZipFile(zipfile_path, 'r')
    file_name = os.path.basename(zipfile_path)  # 获取文件名
    zipdir = os.path.join(os.path.dirname(zipfile_path), str(file_name.split('.')[0]))  # 获取文件所在目录
    for files in file_zip.namelist():
        file_zip.extract(files, os.path.join(zipfile_path, zipdir))  # 解压到指定文件目录
    file_zip.close()
    return True
```

3、读取解压后的文件夹，打印图片路径

```python
def read_img(zipfile_path):
    img_dict = dict()
    if not isfile_exist(zipfile_path):
        return False
    dir_path = os.path.dirname(zipfile_path)  # 获取文件所在目录
    file_name = os.path.basename(zipfile_path)  # 获取文件名
    pic_dir = 'xl' + os.sep + 'media'  # excel变成压缩包后，再解压，图片在media目录
    pic_path = os.path.join(dir_path, str(file_name.split('.')[0]), pic_dir)

    file_list = os.listdir(pic_path)
    for file in file_list:
        filepath = os.path.join(pic_path, file)
        print(filepath)
        img_index = int(re.findall(r'image(\d+)\.', filepath)[0])
        img_dict[img_index] = dict(img_index=img_index, img_path=filepath)
    return img_dict
```

4、解析xml 文件，获取图片在excel表格中的索引位置信息

```python
def get_img_pos_info(zip_file_path, img_dict, img_feature):
    """解析xml 文件，获取图片在excel表格中的索引位置信息"""
    os.path.dirname(zip_file_path)
    dir_path = os.path.dirname(zip_file_path)  # 获取文件所在目录
    file_name = os.path.basename(zip_file_path)  # 获取文件名
    xml_dir = 'xl' + os.sep + 'drawings' + os.sep + 'drawing1.xml'
    xml_path = os.path.join(dir_path, str(file_name.split('.')[0]), xml_dir)
    image_info_dict = parse_xml(xml_path, img_dict, img_feature=img_feature)  # 解析xml 文件， 返回图片索引位置信息
    return image_info_dict


# 重命名解压获取图片位置，及图片表格索引信息
def get_img_info(excel_file_path, img_feature):
    if img_feature not in ["img_index", "img_path", "img_base64"]:
        raise Exception('图片返回参数错误， ["img_index", "img_path", "img_base64"]')
    zip_file_path = copy_change_file_name(excel_file_path)
    if zip_file_path != '':
        if unzip_file(zip_file_path):
            img_dict = read_img(zip_file_path)
            image_info_dict = get_img_pos_info(zip_file_path, img_dict, img_feature)
            return image_info_dict
    return dict()


# 解析xml文件并获取对应图片位置
def parse_xml(file_name, img_dict, img_feature='img_path'):
    # 得到文档对象
    image_info = dict()
    dom_obj = xmldom.parse(file_name)
    # 得到元素对象
    element = dom_obj.documentElement
    def _f(subElementObj):
        for anchor in subElementObj:
            xdr_from = anchor.getElementsByTagName('xdr:from')[0]
            col = xdr_from.childNodes[0].firstChild.data  # 获取标签间的数据
            row = xdr_from.childNodes[2].firstChild.data
            embed = \
            anchor.getElementsByTagName('xdr:pic')[0].getElementsByTagName('xdr:blipFill')[0].getElementsByTagName('a:blip')[0].getAttribute('r:embed')  # 获取属性
            image_info[(int(row), int(col))] = img_dict.get(int(embed.replace('rId', '')), {}).get(img_feature)

    sub_twoCellAnchor = element.getElementsByTagName("xdr:twoCellAnchor")
    sub_oneCellAnchor = element.getElementsByTagName("xdr:oneCellAnchor")
    _f(sub_twoCellAnchor)
    _f(sub_oneCellAnchor)
    return image_info
```

5、读取包含图片的excel数据

```python
def read_excel_info(file_path, img_col_index, img_feature='img_path'):
    img_info_dict = get_img_info(file_path, img_feature)
    book = xlrd.open_workbook(file_path)
    sheet = book.sheet_by_index(0)
    head = dict()
    for i, v in enumerate(sheet.row(0)):
        head[i] = v.value
    info_list = []
    for row_num in range(sheet.nrows):
        d = dict()
        for col_num in range(sheet.ncols):
            if row_num == 0:
                continue
            if 'empty:' in sheet.cell(row_num, col_num).__str__():
                if col_num in img_col_index:
                    d[head[col_num]] = img_info_dict.get((row_num, col_num))
                else:
                    d[head[col_num]] = sheet.cell(row_num, col_num).value
            else:
                d[head[col_num]] = sheet.cell(row_num, col_num).value
        if d:
            info_list.append(d)
    return info_list
```

```
import re
import xml.dom.minidom as xmldom
import os
import zipfile
import shutil
import xlrd

if __name__ == '__main__':
	path = r'E:\\员工信息.xlsx'
	data = read_excel_info(path, img_col_index=[4])
	print(data)
```



