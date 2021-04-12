# ZeusExcel

## 一、前言

#### EasyExcel

EasyExcel是一个基于Java的简单、省内存的读写Excel的开源项目。在尽可能节约内存的情况下支持读写百M的Excel。

64M内存1分钟内读取75M(46W行25列)的Excel

![](https://cdn.nlark.com/yuque/0/2020/png/553000/1584449315232-b7852eec-dd8d-49ee-8880-5fd52eafed30.png)

文档地址：https://www.yuque.com/easyexcel/doc/easyexcel

## 二、ZeusExcel作用

**ZeusExcel** 基于开源的 **EasyExcel**，通过二次封装使之更适合于业务，可实现对表头和数据的校验、非常方便的在Excel中生成错误信息、创建带下拉框的Excel模板、支持Hibernate注解校验、支持自定义表头、自定义表头样式等。

### 1.读Excel

读取Excel的核心便是 `ExcelReadListener<T>` ，该类为抽象类，实现了 **EasyExcel** 的 `ReadListener<T>`接口，**非线程安全**。

`ExcelReadListener` 有四个抽象方法：

* **headCheck**：校验表头信息
* **verify**：校验从Excel中读取到的数据
* **dataHandle**：处理从Excel中读取到的数据
* **doAfterAllDataHandle**：该方法在所有数据处理完成后执行

`ExcelReadListener` 有如下属性：

* **dataList**：从Excel读取到的数据会存入该属性中；
* **batchHandleNum**：一次处理的数据数量，默认值500。EasyExcel加载数据是逐行加载的，每加载一行就会保存到 `dataList` 中，当达到该属性值时，将会执行 `verify` 和 `dataHandle` 方法；
* **errorInfoList**：记录单元格中数据的错误信息，数据全部读取完毕后可通过该list将错误信息写回到Excel中；

* **enabledAnnotationValidation**：是否开启hibernate注解校验，true 开启、false 关闭，默认开启，当开启时每从Excel加载一行数据就会对其进行校验，并将错误信息保存到 `errorInfoList` 中；
* **lastHandleData**：默认值false，若为true`batchHandleNum` 属性将失效，只会在将Excel中的全部数据都加载到 `dataList` 后，才会执行 `verify` 和 `dataHandle` 方法；
* **headErrorMsg**：表头错误信息，当调用 `headCheck` 方法校验表头失败时，应将错误信息写入该属性中。若该属性不为空或空串时，将不再对数据进行解析，流程将会中断。

`AbstractExcelReadListener<T>` 抽象类对 `ExcelReadListener<T>` 的 `headCheck`、`verify` 和 `doAfterAllDataHandle` 三个抽象方法进行了空实现，在只需读取数据处理无其他需求的情况下可直接实现该类即可。

### 2.写Excel

#### 下拉框处理器 DropDownBoxSheetHandler

只需将下拉框的位置和需显示的内容通过构造函数传入或 `addDropDownBoxInfo` 方法添加到 `validationInfoList` 属性中即可。

该属性中的数据类型为 `DropDownBoxInfo`

#### 错误信息处理器 ErrorInfoCommentHandler

该处理器会将错误信息以批注的形式写入Excel，并将其前景色设置为红色。

#### 表头样式处理器 HeadStyleHandler

该处理器修改表头的样式，默认为黑体、18号字、无边框、垂直和水平居中、支持自适应列宽。

自适应列宽可关闭。

该类中的 `multiRowHeadCellStyles` 属性用来保存表头的样式，该属性的类型为 `List<List<CellStyleProperty>>` ，外层list下标对应行索引，内层下标对应列索引。

### 三、快速开始

可通过 `ExcelUtils` 工具类进行快速操作。

实例代码请参考 `ExcelTest`