package com.jz.zeus.excel.write.builder;

import cn.hutool.core.util.StrUtil;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.util.IoUtils;
import com.alibaba.excel.write.builder.ExcelWriterBuilder;
import lombok.SneakyThrows;

import java.io.*;
import java.util.Objects;

/**
 * @Author JZ
 * @Date 2021/4/15 11:38
 */
public class ZeusExcelWriterBuilder {

    private String excelPath;

    private File excelFile;

    private OutputStream excelOutputStream;

    private String templateExcelPath;

    private File templateFile;

    private InputStream templateInputStream;

    public ZeusExcelWriterBuilder() {}

    public ZeusExcelWriterBuilder(String excelPath) {
        this.excelPath = excelPath;
    }

    public ZeusExcelWriterBuilder(File excelFile) {
        this.excelFile = excelFile;
    }

    public ZeusExcelWriterBuilder(OutputStream excelOutputStream) {
        this.excelOutputStream = excelOutputStream;
    }


    public ZeusExcelWriterBuilder withTemplate(String templateExcelPath) {
        this.templateExcelPath = templateExcelPath;
        return this;
    }

    public ZeusExcelWriterBuilder withTemplate(File templateFile) {
        this.templateFile = templateFile;
        return this;
    }

    public ZeusExcelWriterBuilder withTemplate(InputStream templateInputStream) {
        this.templateInputStream = templateInputStream;
        return this;
    }

    public ZeusExcelWriterSheetBuilder sheet() {
        return new ZeusExcelWriterSheetBuilder(build(), null, null);
    }

    public ZeusExcelWriterSheetBuilder sheet(Integer sheetIndex) {
        return new ZeusExcelWriterSheetBuilder(build(), sheetIndex, null);
    }

    public ZeusExcelWriterSheetBuilder sheet(String sheetName) {
        return new ZeusExcelWriterSheetBuilder(build(), null, sheetName);
    }

    public ZeusExcelWriterSheetBuilder sheet(Integer sheetIndex, String sheetName) {
        return new ZeusExcelWriterSheetBuilder(build(), sheetIndex, sheetName);
    }

    @SneakyThrows
    public ZeusExcelWriter build() {
        ExcelWriterBuilder writerBuilder = new ExcelWriterBuilder();
        writerBuilder.useDefaultStyle(false);
        if (StrUtil.isNotBlank(excelPath)) {
            writerBuilder.file(excelPath);
        } else if (excelFile != null) {
            writerBuilder.file(excelFile);
        } else if (excelOutputStream != null) {
            writerBuilder.file(excelOutputStream);
        }
        if (StrUtil.isNotBlank(templateExcelPath)) {
            if (Objects.equals(templateExcelPath, excelPath)) {
                byte[] excelBytes = IoUtils.toByteArray(new FileInputStream(templateExcelPath));
                InputStream inputStream = new ByteArrayInputStream(excelBytes);
                writerBuilder.withTemplate(inputStream);
            } else {
                writerBuilder.withTemplate(templateExcelPath);
            }
        } else if (templateFile != null) {
            String templatePath = templateFile.getPath();
            if (Objects.equals(excelFile.getPath(), templatePath)) {
                byte[] excelBytes = IoUtils.toByteArray(new FileInputStream(templateFile));
                InputStream inputStream = new ByteArrayInputStream(excelBytes);
                writerBuilder.withTemplate(inputStream);
            } else if (excelFile != null && Objects.equals(excelFile.getPath(), templatePath)) {
                byte[] excelBytes = IoUtils.toByteArray(new FileInputStream(excelFile));
                InputStream inputStream = new ByteArrayInputStream(excelBytes);
                writerBuilder.withTemplate(inputStream);
            } else {
                writerBuilder.withTemplate(templateFile);
            }
        } else if (templateInputStream != null) {
            writerBuilder.withTemplate(templateInputStream);
        }
        return new ZeusExcelWriter(writerBuilder.build());
    }

}
