# ZeusExcel

**ZeusExcel** 基于开源的 **EasyExcel**，通过二次封装使之更适合于业务，可实现对表头和数据的校验、非常方便的在Excel中生成错误信息、创建带下拉框（包含级联下拉）的Excel模板、支持扩展表头、动态表头、自定义表头样式等。

`EasyExcel` 文档地址：https://www.yuque.com/easyexcel/doc/easyexcel

# 快速开始

## 1.注解

`ZeusExcel` 在支持 `EasyExcel` 原生注解的同时扩展了一些注解，以支持扩展的功能。

### @ExtendColumn

`@ExtendColumn` 注解用于扩展表头，该注解应与 `@ExcelIgnore` 注解同时使用。

### @HeadColor

该注解用于解决 `EasyExcel` 在使用注解定义单元格颜色时只能使用 `IndexedColors` 的 `index` 选择背景色，无法使用 `RGB` 来生成想要的背景色的问题。

| 属性                    | 类型   | 默认值 | 描述                               |
| ----------------------- | ------ | ------ | ---------------------------------- |
| cellFillForegroundColor | String | ""     | 前景色RGB十六进制，例如："#8EA9DB" |
| cellFillBackgroundColor | String | ""     | 背景色RGB十六进制，例如："#8EA9DB" |

### @ValidationData

| 属性              | 类型     | 默认值 | 描述                                                   |
| ----------------- | -------- | ------ | ------------------------------------------------------ |
| asDicSheet        | boolean  | false  | 是否将下拉框数据作为字典表，作为字典表后将不再隐藏     |
| sheetName         | String   | ""     | 下拉框数据所在sheet的名称                              |
| dicTitle          | String   | ""     | 下拉框数据作为字典表时，表头的提示信息，为默认值时无   |
| options           | String[] |        | 下拉框中的选项                                         |
| rowNum            | int      | 10000  | 需要填充下拉框的行数                                   |
| checkDatavalidity | boolean  | true   | 是否校验单元格中数据的数据属于下拉框中的数据，默认校验 |
| errorTitle        | String   | ""     | 自定义错误box的标题                                    |
| errorMsg          | String   | ""     | 自定义错误box的错误提示信息                            |

## 2.示例

### 定义Excel模板

