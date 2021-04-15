package com.jz.zeus.excel.util;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.util.ArrayUtil;
import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.util.IoUtils;
import com.jz.zeus.excel.CellErrorInfo;
import com.jz.zeus.excel.read.listener.ExcelReadListener;
import com.jz.zeus.excel.read.listener.NoModelReadListener;
import com.jz.zeus.excel.write.handler.ErrorInfoHandler;
import com.jz.zeus.excel.write.handler.ExtendColumnHandler;
import lombok.SneakyThrows;
import lombok.experimental.UtilityClass;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

import java.io.*;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @Author JZ
 * @Date 2021/3/26 11:52
 */
@UtilityClass
public class ExcelUtils {

    /**
     * 读取Excel文件，并添加错误信息
     * @param readListener                  Excel数据读取监听器
     * @param excelName                     Excel全路径名称
     * @param sheetName                     需要读取的sheet名称
     * @param headClass                     表头信息类，也即保存读取到的数据的类
     * @param <T>
     */
    @SneakyThrows
    public <T> void readAndWriteErrorMsg(ExcelReadListener<T> readListener, String excelName,
                                         String sheetName, Class<T> headClass) {
        byte[] excelBytes = IoUtils.toByteArray(new FileInputStream(excelName));
        readAndWriteErrorMsg(readListener, excelBytes, new FileOutputStream(excelName), sheetName, headClass);
    }

    /**
     * 读取Excel文件，并添加错误信息
     * @param readListener                  Excel数据读取监听器
     * @param excelBytes                    Excel 字节数组
     * @param sheetName                     需要读取的sheet名称
     * @param outputStream                  存放错误结果的Excel的输出流
     * @param headClass                     表头信息类，也即保存读取到的数据的类
     * @param <T>
     */
    @SneakyThrows
    public <T> void readAndWriteErrorMsg(ExcelReadListener<T> readListener, byte[] excelBytes, OutputStream outputStream,
                                         String sheetName, Class<T> headClass) {
        EasyExcel.read(new ByteArrayInputStream(excelBytes))
                .sheet(sheetName)
                .head(headClass)
                .registerReadListener(readListener)
                .doRead();
        addErrorInfo(outputStream, excelBytes, sheetName, readListener.getReadAfterHeadRowNum(), readListener.getErrorInfoList());
    }

    /**
     * 读取具有module的Excel
     * @param readListener          Excel读取监听器
     * @param excelName             Excel全路径名
     * @param sheetName             需要读取的sheet名称
     * @param headClass             Excel的module
     * @param <T>
     */
    @SneakyThrows
    public <T> void read(ExcelReadListener<T> readListener, String excelName, String sheetName, Class<T> headClass) {
        read(readListener, new FileInputStream(excelName), sheetName, headClass);
    }

    /**
     * 读取具有module的Excel
     * @param readListener          Excel读取监听器
     * @param inputStream           Excel输入流
     * @param sheetName            需要读取的sheet名称
     * @param headClass            Excel的module
     * @param <T>
     */
    public <T> void read(ExcelReadListener<T> readListener, InputStream inputStream, String sheetName, Class<T> headClass) {
        EasyExcel.read(inputStream)
                .sheet(sheetName)
                .head(headClass)
                .registerReadListener(readListener)
                .doRead();
    }

    /**
     * 当没有module时，读取Excel内容
     * @param readListener        Excel读取监听器
     * @param excelName           Excel全路径名
     * @param sheetName           需要读取的sheet名称
     * @param headRowNum          表头行数，小于等于 0 时表示没有表头
     */
    @SneakyThrows
    public void read(NoModelReadListener readListener, String excelName, String sheetName, Integer headRowNum) {
        read(readListener, new FileInputStream(excelName), sheetName, headRowNum);
    }

    /**
     * 当没有module时，读取Excel内容
     * @param readListener        Excel读取监听器
     * @param inputStream         Excel输入流
     * @param sheetName           需要读取的sheet名称
     * @param headRowNum          表头行数，小于等于 0 时表示没有表头
     */
    public void read(NoModelReadListener readListener, InputStream inputStream, String sheetName, Integer headRowNum) {
        EasyExcel.read(inputStream)
                .sheet(sheetName)
                .headRowNumber(headRowNum)
                .registerReadListener(readListener)
                .doRead();
    }

    /**
     * @param excelName            需要写入错误信息的Excel全路径名称
     * @param sheetName            需要写入错误信息的sheet名称
     * @param headRowNum           表头行数，为null默认1行
     * @param errorInfos           单元格错误信息
     */
    @SneakyThrows
    public void addErrorInfo(String excelName, String sheetName, Integer headRowNum, List<CellErrorInfo> errorInfos) {
        addErrorInfo(excelName, excelName, sheetName, headRowNum, errorInfos);
    }

    /**
     *
     * @param resultExcelName      添加错误信息后保存到的Excel全路径名称
     * @param sourceExcelName      需要写入错误信息的Excel全路径名称
     * @param sheetName            需要写入错误信息的sheet名称
     * @param headRowNum           表头行数，为null默认1行
     * @param errorInfos           单元格错误信息
     */
    @SneakyThrows
    public void addErrorInfo(String resultExcelName, String sourceExcelName, String sheetName, Integer headRowNum, List<CellErrorInfo> errorInfos) {
        if (resultExcelName.equals(sourceExcelName)) {
            byte[] excelBytes = IoUtils.toByteArray(new FileInputStream(sourceExcelName));
            addErrorInfo(new FileOutputStream(resultExcelName), excelBytes, sheetName, headRowNum, errorInfos);
        } else {
            addErrorInfo(new FileOutputStream(resultExcelName), new FileInputStream(sourceExcelName), sheetName, headRowNum, errorInfos);
        }
    }

    /**
     * 如果输入和输出流执行的是同一个文件，会造成文件为空，且会抛出异常，
     * 针对同一个文件建议输入流采用使用字节数组保存Excel，以便流的重复使用 {@link #addErrorInfo}
     * @param resultOutputStream    要写入的Excel的输出流
     * @param sourceInputStream     模板Excel的输入流
     * @param sheetName             需要添加错误信息的sheet名称
     * @param headRowNum            表头行数，为null默认1行
     * @param errorInfos            错误信息
     */
    @SneakyThrows
    public void addErrorInfo(OutputStream resultOutputStream, InputStream sourceInputStream, String sheetName,
                             Integer headRowNum, List<CellErrorInfo> errorInfos) {
        EasyExcel.write(resultOutputStream)
                .withTemplate(sourceInputStream)
                .sheet(sheetName)
                .registerWriteHandler(new ExtendColumnHandler(Collections.emptyList(), null))
                .registerWriteHandler(new ErrorInfoHandler(headRowNum, errorInfos))
                .doWrite(Collections.emptyList());
    }

    /**
     * @param resultOutputStream    要写入的Excel的输出流
     * @param sourceExcelBytes      需要添加错误信息的Excel的字节数组
     * @param sheetName             需要添加错误信息的sheet名称
     * @param headRowNum            表头行数，为null默认1行
     * @param errorInfos            错误信息
     */
    @SneakyThrows
    public void addErrorInfo(OutputStream resultOutputStream, byte[] sourceExcelBytes, String sheetName,
                             Integer headRowNum, List<CellErrorInfo> errorInfos) {
        EasyExcel.write(resultOutputStream)
                .withTemplate(new ByteArrayInputStream(sourceExcelBytes))
                .sheet(sheetName)
                .registerWriteHandler(new ErrorInfoHandler(headRowNum, errorInfos))
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
    public void setCommentErrorInfo(Sheet sheet, Integer rowIndex, Integer columnIndex,
                                    String errorMsgPrefix, String errorMsgSuffix, String... errorMessages) {
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

    /**
     * 获取表头与列索引映射关系
     * @param headMap
     * @return         key 表头、value 列索引
     */
    public Map<String, Integer> getHeadIndexMap(Map<Integer, Head> headMap) {
        if (CollUtil.isEmpty(headMap)) {
            return new HashMap(0);
        }
        Map<String, Integer> headNameIndexMap = new HashMap<>(headMap.size());
        headMap.values().forEach(head -> {
            Integer columnIndex = head.getColumnIndex();
            head.getHeadNameList().forEach(headName -> {
                headNameIndexMap.put(headName, columnIndex);
            });
        });
        return headNameIndexMap;
    }

}
