package com.jz.zeus.excel.util;

import cn.hutool.core.util.ArrayUtil;
import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.util.IoUtils;
import com.jz.zeus.excel.CellErrorInfo;
import com.jz.zeus.excel.DropDownBoxInfo;
import com.jz.zeus.excel.read.listener.AbstractExcelReadListener;
import com.jz.zeus.excel.write.handler.DefaultHeadStyleHandler;
import com.jz.zeus.excel.write.handler.DropDownBoxSheetHandler;
import com.jz.zeus.excel.write.handler.ErrorInfoCommentHandler;
import lombok.SneakyThrows;
import lombok.experimental.UtilityClass;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

import java.io.*;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

/**
 * @Author JZ
 * @Date 2021/3/26 11:52
 */
@UtilityClass
public class ExcelUtils {

    /**
     * 创建一个有表头的空Excel
     * @param outputStream             Excel的输出流
     * @param sheetName                sheet名，可以为null
     * @param headList                 表头
     * @param dropDownBoxInfoList      下拉框配置信息
     * @param <T>
     */
    public <T> void createTemplate(OutputStream outputStream, String sheetName, List<String> headList, List<DropDownBoxInfo> dropDownBoxInfoList) {
        List<List<String>> heads = new ArrayList<>(1);
        heads.add(headList);
        EasyExcel.write(outputStream)
                .sheet(sheetName)
                .head(heads)
                .registerWriteHandler(new DefaultHeadStyleHandler())
                .registerWriteHandler(new DropDownBoxSheetHandler(dropDownBoxInfoList))
                .doWrite(Collections.emptyList());
    }

    /**
     * 创建一个有表头的空Excel
     * @param outputStream             Excel的输出流
     * @param sheetName                sheet名，可以为null
     * @param dataClass                表示表头信息的class
     * @param dropDownBoxInfoList      下拉框配置信息
     * @param excludeColumnFiledNames   不需要包含的表头，可以为null
     * @param <T>
     */
    public <T> void createTemplate(OutputStream outputStream, String sheetName, Class<T> dataClass, List<DropDownBoxInfo> dropDownBoxInfoList, List<String> excludeColumnFiledNames) {
        createTemplate(outputStream, sheetName, dataClass, null, dropDownBoxInfoList, excludeColumnFiledNames);
    }

    /**
     * 创建一个有表头的空Excel
     * @param outputStream             Excel的输出流
     * @param sheetName                sheet名，可以为null
     * @param dataClass                表示表头信息的class
     * @param headRowNum               表示表头行数
     * @param dropDownBoxInfoList      下拉框配置信息
     * @param excludeColumnFiledNames   不需要包含的表头，可以为null
     * @param <T>
     */
    public <T> void createTemplate(OutputStream outputStream, String sheetName, Class<T> dataClass, Integer headRowNum, List<DropDownBoxInfo> dropDownBoxInfoList, List<String> excludeColumnFiledNames) {
        EasyExcel.write(outputStream)
                .sheet(sheetName)
                .head(dataClass)
                .registerWriteHandler(new DefaultHeadStyleHandler())
                .registerWriteHandler(new DropDownBoxSheetHandler(headRowNum, dropDownBoxInfoList))
                .excludeColumnFiledNames(excludeColumnFiledNames)
                .doWrite(Collections.emptyList());
    }

    /**
     * 创建一个有表头的空Excel
     * @param excelName                Excel全路径名
     * @param sheetName                sheet名，可以为null
     * @param headList                 表头
     * @param dropDownBoxInfoList      下拉框配置信息
     * @param <T>
     */
    public <T> void createTemplate(String excelName, String sheetName, List<String> headList, List<DropDownBoxInfo> dropDownBoxInfoList) {
        List<List<String>> heads = new ArrayList<>(1);
        heads.add(headList);
        EasyExcel.write(excelName)
                .sheet(sheetName)
                .head(heads)
                .registerWriteHandler(new DefaultHeadStyleHandler())
                .registerWriteHandler(new DropDownBoxSheetHandler(dropDownBoxInfoList))
                .doWrite(Collections.emptyList());
    }

    /**
     * 创建一个有表头的空Excel
     * @param excelName                Excel全路径名
     * @param sheetName                sheet名，可以为null
     * @param dataClass                表示表头信息的class
     * @param dropDownBoxInfoList      下拉框配置信息
     * @param excludeColumnFiledNames   不需要包含的表头，可以为null
     * @param <T>
     */
    public <T> void createTemplate(String excelName, String sheetName, Class<T> dataClass, List<DropDownBoxInfo> dropDownBoxInfoList, List<String> excludeColumnFiledNames) {
        createTemplate(excelName, sheetName, dataClass, null, dropDownBoxInfoList, excludeColumnFiledNames);
    }

    /**
     * 创建一个有表头的空Excel
     * @param excelName                Excel全路径名
     * @param sheetName                sheet名，可以为null
     * @param dataClass                表示表头信息的class
     * @param headRowNum               表示表头行数
     * @param dropDownBoxInfoList      下拉框配置信息
     * @param excludeColumnFiledNames   不需要包含的表头，可以为null
     * @param <T>
     */
    public <T> void createTemplate(String excelName, String sheetName, Class<T> dataClass, Integer headRowNum, List<DropDownBoxInfo> dropDownBoxInfoList, List<String> excludeColumnFiledNames) {
        EasyExcel.write(excelName)
                .sheet(sheetName)
                .head(dataClass)
                .registerWriteHandler(new DefaultHeadStyleHandler())
                .registerWriteHandler(new DropDownBoxSheetHandler(headRowNum, dropDownBoxInfoList))
                .excludeColumnFiledNames(excludeColumnFiledNames)
                .doWrite(Collections.emptyList());
    }

    /**
     * 读取Excel文件，并添加错误信息
     */
    @SneakyThrows
    public <T> void readAndWriteErrorMsg(AbstractExcelReadListener<T> readListener, String excelName, String sheetName, Class<T> dataClass) {
        byte[] excelBytes = IoUtils.toByteArray(new FileInputStream(excelName));
        readAndWriteErrorMsg(readListener, excelBytes, new FileOutputStream(excelName), sheetName, dataClass);
    }

    /**
     * 读取Excel字节数组，并添加错误信息
     */
    @SneakyThrows
    public <T> void readAndWriteErrorMsg(AbstractExcelReadListener<T> readListener, byte[] excelBytes, OutputStream outputStream, String sheetName, Class<T> dataClass) {
        EasyExcel.read(new ByteArrayInputStream(excelBytes))
                .sheet(sheetName)
                .head(dataClass)
                .registerReadListener(readListener)
                .doRead();
        if (!readListener.hasDataError()) {
            addErrorInfo(outputStream, excelBytes, sheetName, readListener.getErrorInfoList());
        }
    }

    /**
     *
     * @param resultExcelName
     * @param templateExcelName
     * @param errorInfos
     */
    @SneakyThrows
    public void addErrorInfo(String resultExcelName, String templateExcelName, String sheetName, List<CellErrorInfo> errorInfos) {
        if (resultExcelName.equals(templateExcelName)) {
            byte[] excelBytes = IoUtils.toByteArray(new FileInputStream(templateExcelName));
            addErrorInfo(new FileOutputStream(resultExcelName), excelBytes, sheetName, errorInfos);
        } else {
            addErrorInfo(new FileOutputStream(resultExcelName), new FileInputStream(templateExcelName), sheetName, errorInfos);
        }
    }

    /**
     * @param resultOutputStream    要写入的Excel的输出流
     * @param templateInputStream   模板Excel的输入流
     * @param errorInfos            错误信息
     */
    @SneakyThrows
    public void addErrorInfo(OutputStream resultOutputStream, InputStream templateInputStream, String sheetName, List<CellErrorInfo> errorInfos) {
        EasyExcel.write(resultOutputStream)
                .withTemplate(templateInputStream)
                .sheet(sheetName)
                .registerWriteHandler(new ErrorInfoCommentHandler(errorInfos))
                .doWrite(Collections.emptyList());
    }

    /**
     * @param resultOutputStream    要写入的Excel的输出流
     * @param templateExcelBytes    模板Excel的字节数组
     * @param errorInfos            错误信息
     */
    @SneakyThrows
    public void addErrorInfo(OutputStream resultOutputStream, byte[] templateExcelBytes, String sheetName, List<CellErrorInfo> errorInfos) {
        EasyExcel.write(resultOutputStream)
                .withTemplate(new ByteArrayInputStream(templateExcelBytes))
                .sheet(sheetName)
                .registerWriteHandler(new ErrorInfoCommentHandler(errorInfos))
                .doWrite(Collections.emptyList());
    }

    public void setCommentErrorInfo(Sheet sheet, Integer rowIndex, Integer columnIndex, String... errorMessages) {
        setCommentErrorInfo(sheet, rowIndex, columnIndex, "- ", "", errorMessages);
    }

    /**
     * 给sheet的指定单元格设置错误信息，错误信息将显示在批注里，且单元格背景色为红色
     * @param sheet                 单元格 所在sheet
     * @param rowIndex              单元格 的行索引
     * @param columnIndex           单元格 的列索引
     * @param errorMsgPrefix        错误信息前缀
     * @param errorMsgSuffix        错误信息后缀
     * @param errorMessages         错误信息
     */
    public void setCommentErrorInfo(Sheet sheet, Integer rowIndex, Integer columnIndex, String errorMsgPrefix, String errorMsgSuffix, String... errorMessages) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            return;
        }
        Cell cell = row.getCell(columnIndex);
        if (cell == null) {
            return;
        }
        CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellStyle.setFillForegroundColor(IndexedColors.RED.index);
        cell.setCellStyle(cellStyle);

        Drawing<?> drawing = sheet.createDrawingPatriarch();
        Comment comment = drawing.createCellComment(new XSSFClientAnchor(0, 0, 0, 0, columnIndex, rowIndex, columnIndex+2, rowIndex+2));
        comment.setString(new XSSFRichTextString(ArrayUtil.join(errorMessages, "\n", errorMsgPrefix, errorMsgSuffix)));
        cell.setCellComment(comment);
    }

}
