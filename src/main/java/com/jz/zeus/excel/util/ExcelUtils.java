package com.jz.zeus.excel.util;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.util.ArrayUtil;
import cn.hutool.core.util.StrUtil;
import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.util.ClassUtils;
import com.alibaba.excel.util.IoUtils;
import com.jz.zeus.excel.CellErrorInfo;
import com.jz.zeus.excel.DropDownBoxInfo;
import com.jz.zeus.excel.read.listener.AbstractExcelReadListener;
import com.jz.zeus.excel.write.handler.DropDownBoxSheetHandler;
import com.jz.zeus.excel.write.handler.ErrorInfoCommentHandler;
import com.jz.zeus.excel.write.handler.HeadStyleHandler;
import lombok.SneakyThrows;
import lombok.experimental.UtilityClass;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.hibernate.validator.internal.util.ConcurrentReferenceHashMap;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.lang.reflect.Field;
import java.net.URLEncoder;
import java.util.*;

/**
 * @Author JZ
 * @Date 2021/3/26 11:52
 */
@UtilityClass
public class ExcelUtils {

    private static final Map<Class, List<List<String>>> CLASS_HEAD_CACHE_MAP = new ConcurrentReferenceHashMap<>();

    /**
     * 将生成的Excel直接写入response
     * @param response
     * @param excelName                 Excel名
     * @param sheetName                 sheet名，可以为 null
     * @param headClass                 表示表头信息的 class
     * @param dropDownBoxInfoList       下拉框配置信息
     * @param excludeColumnFiledNames   不需要包含的表头，可以为 null
     */
    @SneakyThrows
    public void downloadTemplate(HttpServletResponse response, String excelName, String sheetName, Class<?> headClass, HeadStyleHandler headStyleHandler, List<DropDownBoxInfo> dropDownBoxInfoList, List<String> excludeColumnFiledNames) {
        response.setContentType("application/vnd.ms-excel");
        response.setCharacterEncoding("utf-8");
        // 这里URLEncoder.encode可以防止中文乱码
        String fileName = URLEncoder.encode(excelName, "UTF-8").replaceAll("\\+", "%20");
        response.setHeader("Content-disposition", "attachment;filename*=utf-8''" + fileName + ".xlsx");
        createTemplate(response.getOutputStream(), sheetName, headClass, headStyleHandler, dropDownBoxInfoList, excludeColumnFiledNames);
    }

    /**
     * 将生成的Excel直接写入response
     * @param response
     * @param excelName                Excel名
     * @param sheetName                sheet名，可以为null
     * @param headList                 表头
     * @param dropDownBoxInfoList      下拉框配置信息
     */
    @SneakyThrows
    public void downloadTemplate(HttpServletResponse response, String excelName, String sheetName, List<String> headList, HeadStyleHandler headStyleHandler, List<DropDownBoxInfo> dropDownBoxInfoList) {
        response.setContentType("application/vnd.ms-excel");
        response.setCharacterEncoding("utf-8");
        // 这里URLEncoder.encode可以防止中文乱码
        String fileName = URLEncoder.encode(excelName, "UTF-8").replaceAll("\\+", "%20");
        response.setHeader("Content-disposition", "attachment;filename*=utf-8''" + fileName + ".xlsx");
        createTemplate(response.getOutputStream(), sheetName, headList, headStyleHandler, dropDownBoxInfoList);
    }

    /**
     * 创建一个有表头的空Excel
     * @param outputStream             Excel的输出流
     * @param sheetName                sheet名，可以为null
     * @param headList                 表头
     * @param dropDownBoxInfoList      下拉框配置信息
     */
    public void createTemplate(OutputStream outputStream, String sheetName, List<String> headList, HeadStyleHandler headStyleHandler, List<DropDownBoxInfo> dropDownBoxInfoList) {
        List<List<String>> heads = new ArrayList<>(headList.size());
        headList.forEach(head -> heads.add(new ArrayList<String>(1) {{add(head);}}));
        EasyExcel.write(outputStream)
                .useDefaultStyle(false)
                .sheet(sheetName)
                .head(heads)
                .registerWriteHandler(headStyleHandler == null ? new HeadStyleHandler() : headStyleHandler)
                .registerWriteHandler(new DropDownBoxSheetHandler(dropDownBoxInfoList))
                .doWrite(Collections.emptyList());
    }

    /**
     * 创建一个有表头的空Excel
     * @param outputStream              Excel的输出流
     * @param sheetName                 sheet名，可以为 null
     * @param headClass                 表示表头信息的 class
     * @param dropDownBoxInfoList       下拉框配置信息
     * @param excludeColumnFiledNames   不需要包含的表头，可以为 null
     */
    public void createTemplate(OutputStream outputStream, String sheetName, Class<?> headClass, HeadStyleHandler headStyleHandler, List<DropDownBoxInfo> dropDownBoxInfoList, List<String> excludeColumnFiledNames) {
        EasyExcel.write(outputStream)
                .useDefaultStyle(false)
                .sheet(sheetName)
                .head(headClass)
                .registerWriteHandler(headStyleHandler == null ? new HeadStyleHandler() : headStyleHandler)
                .registerWriteHandler(new DropDownBoxSheetHandler(dropDownBoxInfoList))
                .excludeColumnFiledNames(excludeColumnFiledNames)
                .doWrite(Collections.emptyList());
    }

    /**
     * 创建一个有表头的空Excel
     * @param excelName                Excel全路径名
     * @param sheetName                sheet名，可以为null
     * @param headList                 表头
     * @param dropDownBoxInfoList      下拉框配置信息
     */
    public void createTemplate(String excelName, String sheetName, List<String> headList, HeadStyleHandler headStyleHandler, List<DropDownBoxInfo> dropDownBoxInfoList) {
        List<List<String>> heads = new ArrayList<>(headList.size());
        headList.forEach(head -> heads.add(new ArrayList<String>(1) {{add(head);}}));
        EasyExcel.write(excelName)
                .useDefaultStyle(false)
                .sheet(sheetName)
                .head(heads)
                .registerWriteHandler(headStyleHandler == null ? new HeadStyleHandler() : headStyleHandler)
                .registerWriteHandler(new DropDownBoxSheetHandler(dropDownBoxInfoList))
                .doWrite(Collections.emptyList());
    }

    /**
     * 创建一个有表头的空Excel
     * @param excelName                Excel全路径名
     * @param sheetName                sheet名，可以为null
     * @param headClass                表示表头信息的class
     * @param dropDownBoxInfoList      下拉框配置信息
     * @param excludeColumnFiledNames  不需要读取的字段名，字段名对应 dataClass 中的字段名，并非表头名
     */
    public void createTemplate(String excelName, String sheetName, Class<?> headClass, List<DropDownBoxInfo> dropDownBoxInfoList, List<String> excludeColumnFiledNames) {
        EasyExcel.write(excelName)
                .useDefaultStyle(false)
                .sheet(sheetName)
                .head(headClass)
                .registerWriteHandler(new HeadStyleHandler())
                .registerWriteHandler(new DropDownBoxSheetHandler(dropDownBoxInfoList))
                .excludeColumnFiledNames(excludeColumnFiledNames)
                .doWrite(Collections.emptyList());
    }

    /**
     * 读取Excel文件，并添加错误信息
     * @param readListener                  Excel数据读取监听器
     * @param excelName                     Excel全路径名称
     * @param sheetName                     需要读取的sheet名称
     * @param headClass                     表头信息类，也即保存读取到的数据的类
     * @param <T>
     */
    @SneakyThrows
    public <T> void readAndWriteErrorMsg(AbstractExcelReadListener<T> readListener, String excelName, String sheetName, Class<T> headClass, List<String> excludeColumnFiledNames) {
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
    public <T> void readAndWriteErrorMsg(AbstractExcelReadListener<T> readListener, byte[] excelBytes, OutputStream outputStream, String sheetName, Class<T> headClass) {
        EasyExcel.read(new ByteArrayInputStream(excelBytes))
                .sheet(sheetName)
                .head(headClass)
                .registerReadListener(readListener)
                .doRead();
        addErrorInfo(outputStream, excelBytes, sheetName, null, readListener.getErrorInfoList());
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
    public void addErrorInfo(OutputStream resultOutputStream, InputStream sourceInputStream, String sheetName, Integer headRowNum, List<CellErrorInfo> errorInfos) {
        EasyExcel.write(resultOutputStream)
                .withTemplate(sourceInputStream)
                .sheet(sheetName)
                .registerWriteHandler(new ErrorInfoCommentHandler(headRowNum, errorInfos))
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
    public void addErrorInfo(OutputStream resultOutputStream, byte[] sourceExcelBytes, String sheetName, Integer headRowNum, List<CellErrorInfo> errorInfos) {
        EasyExcel.write(resultOutputStream)
                .withTemplate(new ByteArrayInputStream(sourceExcelBytes))
                .sheet(sheetName)
                .registerWriteHandler(new ErrorInfoCommentHandler(headRowNum, errorInfos))
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

    public List<List<String>> getClassHeads(Class<?> clazz) {
        List<List<String>> result = CLASS_HEAD_CACHE_MAP.get(clazz);
        if (result != null) {
            return result;
        }
        Map<Integer, Field> filedMap = new HashMap<>();
        ClassUtils.declaredFields(clazz, filedMap, null, null, true, false, null);
        if (CollUtil.isEmpty(filedMap)) {
            return null;
        }
        result = new ArrayList<>();
        for (Map.Entry<Integer, Field> entry : filedMap.entrySet()) {
            Field field = entry.getValue();
            ExcelProperty excelProperty = field.getAnnotation(ExcelProperty.class);
            if (excelProperty == null) {
                List<String> rowHeads;
                if (result.size() == 0) {
                    rowHeads = new ArrayList<>();
                    result.add(rowHeads);
                } else {
                    rowHeads = result.get(0);
                }
                rowHeads.add(entry.getKey(), field.getName());
                continue;
            }
            String[] headArray = excelProperty.value();
            for (int rowIndex = 0; rowIndex < headArray.length; rowIndex++) {
                if (StrUtil.isBlank(headArray[rowIndex])) {
                    continue;
                }
                List<String> rowHeads;
                if (rowIndex >= result.size()) {
                    rowHeads = new ArrayList<>();
                    result.add(rowHeads);
                } else {
                    rowHeads = result.get(rowIndex);
                }
                rowHeads.add(entry.getKey(), headArray[rowIndex]);
            }
        }
        return result;
    }

}
