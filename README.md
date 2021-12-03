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

另一种方式是使用`openpyxl`和`openpyxl_image_loader`库，按行读取，loader 图片进行保存，完整代码见：[new_read_data.py](https://github.com/dddreams/read-excel-image/blob/master/new_read_data.pyhttps://github.com/dddreams/read-excel-image/blob/master/new_read_data.py)。

## 7、新增的需求

- 循环读取在某个目录下的多个文件))
  
  ```python
  for root, dirs, files in os.walk(source_root):
    for file in files:
      print(os.path.join(root, file))
  ```

- leader 要求照片大于200K不入库，于是添加了压缩图片的功能，我将压缩图片的代码分离了出来[compress_image.py](https://github.com/dddreams/read-excel-image/blob/master/compress_image.py)。

- 将有问题的数据记录下来，写入 Excel，于是有了写入Excel的代码。
  
  ```python
  wb = Workbook()
  ws = wb.create_sheet("存在问题的数据", 0)
  index = 1
  for i in range(len(error_data)):
    index = index + 1
    arr_list = error_data[i].split("|")
    for j in range(len(arr_list)):
      ws.cell(row = index, column= j+1, value = arr_list[j])
  wb.save(target_root + '存在问题的数据.xlsx')
  ```

- 照片使用电话号码命名，并生成日志，写入文件。

## 8、存在的问题

由于原始数据中存在照片未采集的记录，但是提取到的数据中这些记录都有对应的照片，原来`image_loader = SheetImageLoader(ws)`每次读完不会清空字典，所以就会把上一个文件中对应行的照片读取到当前文件的这一行，经过搜索查找发现是`openpyxl-image-loader`的问题，相关`issues`地址：[images should not be static variable of SheetImageLoader](https://github.com/ultr4nerd/openpyxl-image-loader/issues/9) 。所以在每次循环结束将`image_loader` 清空即可，添加这行代码：

```python
image_loader._images.clear()
```

## 9、通过VB导出图片

其实提取Excel中的图片可以使用VB实现，直接在Excel的sheet上右键【查看代码】然后粘贴一下代码执行就会将图片导出来，并且能以任一列的值命名。

```vb
Sub 导出图片()
    On Error Resume Next
    MkDir ThisWorkbook.Path & "\图片"
    For Each pic In ActiveSheet.Shapes
        If pic.Type = 13 Then
            RN = pic.TopLeftCell.Offset(0, -3).Value
            pic.Copy
            With ActiveSheet.ChartObjects.Add(0, 0, pic.Width, pic.Height).Chart    '创建图片
                .Parent.Select
                .Paste
                .Export ThisWorkbook.Path & "\图片\" & RN & ".jpg"
                .Parent.Delete
            End With
        End If
    Next
    MsgBox "导出图片完成！ "
End Sub
```

## 10、总结

经过不断的折腾，发现条条大路通罗马才是真理，不管你用什么方式实现，发现问题、解决问题才是最重要的经历。
