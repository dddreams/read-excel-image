## 1、需求

有这样一个需求，将采集在 Excel 中的人员信息（包含照片）导入到 Mysql 库。需求很简单，就是读取 Excel 中的数据插入到 Mysql 中的表，问题在于照片怎么读取？

## 2、分析

对于读取文本数据，直接按行读取即可；对于图片，常用做法是，将 Excel 文件的后缀名改为 zip，然后解压文件，对应文件名目录下有一个`\xl\media`的目录，里面便是我们要照片的图片，然而，他的文件名、文件顺序都是乱的，确定不出图片对应的记录，怎么办呢？其实目录中还有这样一个 xml 文件`\xl\drawings\drawing1.xml`其中就有图片与记录的对应关系，只要解析 xml 文件就可以了。知道了这个思路，那我们来一步步的实现他。

## 3、实现

完整代码  [https://github.com/dddreams/read-excel-image/blob/master/read_users.py](https://github.com/dddreams/read-excel-image/blob/master/read_users.py)

## 4、图片对应不到的问题

有时会出现图片对应不到记录的情况，这是因为 Excel 中图片不在单独的单元格内，指定的列中取不到图片，这就需要调整 Excel 内的数据了，但是如果数据量比较大，调整起来就比较麻烦了。其实，我们可以借助 Excel 宏的功能，批量修改图片大小，批量让其居中，位于单元格内部，这样在解析 xml 时就能对应到每条记录了。

打开 Excel 按下 `Alt + F11` ，然后点击【插入】菜单，选择【模块】复制下面代码：

```
Sub dq()
Dim shp As Shape
For Each shp In ActiveSheet.Shapes
shp.Left = (shp.TopLeftCell.Width - shp.Width) / 2 + shp.TopLeftCell.Left
shp.Top = (shp.TopLeftCell.Height - shp.Height) / 2 + shp.TopLeftCell.Top
Next
End Sub
```

然后返回 Excel 【开始】菜单中点击【查找和选择】选择【定位条件】中【对象】，将选中所有的图片，然后按下`Alt+F8` 点击【执行】就可以批量将图片在所在单元格居中了。如果图片大小过大，全部选中后更改图片大小即可。然后解析 Xml 文件就不会出现图片与记录对应不到的情况了。

## 5、经过实践后的问题

代码经过实践后，发现还是有问题，有些图片还是对应不到相应的记录，于是又开始了一波debugger，发现不是代码的锅，而是Excel解压后`drawing1.xml`的锅，来看看我们解析xml的代码：

```python
def _f(subElementObj):
    for anchor in subElementObj:
        xdr_from = anchor.getElementsByTagName('xdr:from')[0]
        col = xdr_from.childNodes[0].firstChild.data  # 获取标签间的数据
        row = xdr_from.childNodes[2].firstChild.data
        embed = anchor.getElementsByTagName('xdr:pic')[0].getElementsByTagName('xdr:blipFill')[0].getElementsByTagName('a:blip')[0].getAttribute('r:embed')  # 获取属性
        image_info[(int(row), int(col))] = img_dict.get(int(embed.replace('rId', '')), {}).get(img_feature)
```

要解析的xml文档部分内容：

```xml
<xdr:pic>
  <xdr:blipFill>
    ...
    <a:blip r:embed="rId1" cstate="print">
      ...
    </a:blip>
  </xdr:blipFill>
</xdr:pic>
```

获取到`<a:blip>`元素的`r:embed`属性，即对应团片的序号，实际上，如果Excel内容是从其他地方复制过来的，他的序号与图片的序号对应不上，导致的问题，遗憾的是没找到什么原因，不知道Excel中是如何对应的，有兴趣的同学可以研究下。

## 6、另一种方式的实现

另一种方式是使用`openpyxl`和`openpyxl_image_loader`库，按行读取，loader 图片进行保存，完整代码见：[https://github.com/dddreams/read-excel-image/blob/master/new_read_data.pyhttps://github.com/dddreams/read-excel-image/blob/master/new_read_data.py](https://github.com/dddreams/read-excel-image/blob/master/new_read_data.pyhttps://github.com/dddreams/read-excel-image/blob/master/new_read_data.py)。

## 7、新增的需求




