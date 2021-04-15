package com.jz.zeus.excel;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.read.builder.ExcelReaderBuilder;
import com.jz.zeus.excel.write.builder.ZeusExcelWriterBuilder;
import com.jz.zeus.excel.write.builder.ZeusExcelWriterSheetBuilder;
import lombok.SneakyThrows;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URLEncoder;

/**
 * @Author JZ
 * @Date 2021/4/15 10:18
 */
public class ZeusExcel {

    public static ExcelReaderBuilder read(String excelPath) {
        return EasyExcel.read(excelPath);
    }

    public static ExcelReaderBuilder read(File excelFile) {
        return EasyExcel.read(excelFile);
    }

    public static ExcelReaderBuilder read(InputStream excelInputStream) {
        return EasyExcel.read(excelInputStream);
    }

    public static ZeusExcelWriterBuilder write() {
        return new ZeusExcelWriterBuilder();
    }

    public static ZeusExcelWriterBuilder write(String excelPath) {
        return new ZeusExcelWriterBuilder(excelPath);
    }

    public static ZeusExcelWriterBuilder write(File excelFile) {
        return new ZeusExcelWriterBuilder(excelFile);
    }

    public static ZeusExcelWriterBuilder write(OutputStream excelOutputStream) {
        return new ZeusExcelWriterBuilder(excelOutputStream);
    }

    @SneakyThrows
    public static ZeusExcelWriterBuilder write(HttpServletResponse response, String excelName) {
        response.setContentType("application/vnd.ms-excel");
        response.setCharacterEncoding("utf-8");
        // 这里URLEncoder.encode可以防止中文乱码
        String fileName = URLEncoder.encode(excelName, "UTF-8").replaceAll("\\+", "%20");
        response.setHeader("Content-disposition", "attachment;filename*=utf-8''" + fileName + ".xlsx");
        return write(response.getOutputStream());
    }

    public static ZeusExcelWriterSheetBuilder writeSheet(String sheetName) {
        return new ZeusExcelWriterSheetBuilder(null, sheetName);
    }

    public static ZeusExcelWriterSheetBuilder writeSheet(Integer sheetIndex) {
        return new ZeusExcelWriterSheetBuilder(sheetIndex, null);
    }

    public static ZeusExcelWriterSheetBuilder writeSheet(Integer sheetIndex, String sheetName) {
        return new ZeusExcelWriterSheetBuilder(sheetIndex, sheetName);
    }

}
