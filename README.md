## 1、需求

有这样一个需求，将采集在 Excel 中的人员信息（包含照片）导入到 Mysql 库。需求很简单，就是读取 Excel 中的数据插入到 Mysql 中的表，问题在于照片怎么读取？

## 2、分析

对于读取文本数据，直接按行读取即可；对于图片，常用做法是，将 Excel 文件的后缀名改为 zip，然后解压文件，对应文件名目录下有一个`\xl\media`的目录，里面便是我们要照片的图片，然而，他的文件名、文件顺序都是乱的，确定不出图片对应的记录，怎么办呢？其实目录中还有这样一个 xml 文件`\xl\drawings\drawing1.xml`其中就有图片与记录的对应关系，只要解析 xml 文件就可以了。知道了这个思路，那我们来一步步的实现他。

## 3、实现

具体代码见 read_users.py  [https://github.com/dddreams/read-excel-image/blob/master/read_users.py](https://github.com/dddreams/read-excel-image/blob/master/read_users.py)

## 4、常见问题



