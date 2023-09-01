## 概述
对 [Excelize](https://github.com/qax-os/excelize) 进行了简单的封装，方便的读取表格和往表格里面写入数据，大部分情况下，定义好结构体就行， 而不需要在代码里面指定各种坐标，让代码难以维护和阅读
> 行和列从1开始算

## 字段标签
| 标签名      | 说明 |
| ----------- | ----------- |
| axis      |指定单元格坐标，不指定时，默认为字段在结构体的顺序，如为结构体的第一个字段，则列的坐标为A，第一列|
| style   | 指定单元格样式，excelize.Style{}序列化后的json字符串，参考文档：https://xuri.me/excelize/zh-hans/style.html |
| column   | 指定列名，用于根据表头的和列名来匹配对应的列坐标|
| colWidth   | 指定列的宽度 |

## 示例
### 写入一个结构体数据

### 写入多行数据

### 写入多行数据

### 模板语法

### 读取多行数据

## 参考链接
* Excelize文档：https://xuri.me/excelize/zh-hans/