package com.jz.zeus.excel.write.builder;

import cn.hutool.core.bean.BeanUtil;
import cn.hutool.core.collection.CollUtil;
import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.exception.ExcelGenerateException;
import com.alibaba.excel.write.builder.ExcelWriterSheetBuilder;
import com.jz.zeus.excel.CellErrorInfo;
import com.jz.zeus.excel.DynamicHead;
import com.jz.zeus.excel.ValidationInfo;
import com.jz.zeus.excel.context.ExcelContext;
import com.jz.zeus.excel.write.handler.*;
import com.jz.zeus.excel.write.property.CellStyleProperty;

import java.util.ArrayList;
import java.util.List;

/**
 * @author:JZ
 * @date:2021/4/14
 */
public class ZeusExcelWriterSheetBuilder {

    private static final Integer SHEET_INDEX_DEFAULT = 0;

    private ZeusExcelWriter excelWriter;

    private ExcelContext excelContext = new ExcelContext();

    private Integer sheetIndex;

    private String sheetName;

    private Boolean needHead;

    private Boolean removeOldErrorInfo;

    private List<DynamicHead> dynamicHeads;

    private List<ValidationInfo> validationInfos;

    private List<CellErrorInfo> errorInfos;

    private CellStyleProperty headStyle;

    private List<CellStyleProperty> headStyles;

    private List<String> excludeColumnFiledNames;


    public ZeusExcelWriterSheetBuilder(Integer sheetIndex, String sheetName) {
        this.sheetIndex = sheetIndex == null ? SHEET_INDEX_DEFAULT : sheetIndex;
        this.sheetName = sheetName;
    }

    public ZeusExcelWriterSheetBuilder(ZeusExcelWriter excelWriter, Integer sheetIndex, String sheetName) {
        this.excelWriter = excelWriter;
        this.sheetIndex = sheetIndex == null ? SHEET_INDEX_DEFAULT : sheetIndex;
        this.sheetName = sheetName;
    }

    public ZeusExcelWriterSheetBuilder dynamicHeads(List<DynamicHead> dynamicHeads) {
        this.dynamicHeads = dynamicHeads;
        return this;
    }

    public ZeusExcelWriterSheetBuilder validationInfos(List<ValidationInfo> validationInfos) {
        this.validationInfos = validationInfos;
        return this;
    }

    public ZeusExcelWriterSheetBuilder errorInfos(List<CellErrorInfo> errorInfos) {
        this.errorInfos = errorInfos;
        return this;
    }

    public ZeusExcelWriterSheetBuilder headStyle(CellStyleProperty headStyle) {
        this.headStyle = headStyle;
        return this;
    }

    public ZeusExcelWriterSheetBuilder headStyles(List<CellStyleProperty> headStyles) {
        this.headStyles = headStyles;
        return this;
    }

    public ZeusExcelWriterSheetBuilder excludeColumnFiledNames(List<String> excludeColumnFiledNames) {
        this.excludeColumnFiledNames = excludeColumnFiledNames;
        return this;
    }

    public ZeusExcelWriterSheetBuilder needHead(Boolean needHead) {
        this.needHead = needHead;
        return this;
    }

    public ZeusExcelWriterSheetBuilder removeOldErrorInfo(Boolean removeOldErrorInfo) {
        this.removeOldErrorInfo = removeOldErrorInfo;
        return this;
    }

    public ZeusWriteSheet build(List<List<String>> headNames) {
        ExcelWriterSheetBuilder sheetBuilder = EasyExcel.writerSheet(sheetIndex, sheetName);
        sheetBuilder.head(headNames);
        sheetBuilder.registerWriteHandler(new ValidationInfoHandler(excelContext, validationInfos));
        sheetBuilder.registerWriteHandler(new ErrorInfoHandler(excelContext, removeOldErrorInfo, errorInfos));
        if (CollUtil.isNotEmpty(headStyles)) {
            sheetBuilder.registerWriteHandler(new HeadStyleHandler(excelContext, headStyles));
        } else {
            sheetBuilder.registerWriteHandler(new HeadStyleHandler(excelContext, headStyle));
        }
        sheetBuilder.needHead(needHead);
        ZeusWriteSheet zeusWriteSheet = new ZeusWriteSheet(excelContext);
        BeanUtil.copyProperties(sheetBuilder.build(), zeusWriteSheet);
        return zeusWriteSheet;
    }

    public ZeusWriteSheet build(Class headClass) {
        ExcelWriterSheetBuilder sheetBuilder = EasyExcel.writerSheet(sheetIndex, sheetName);
        sheetBuilder.head(headClass);
        if (CollUtil.isNotEmpty(excludeColumnFiledNames)) {
            sheetBuilder.excludeColumnFiledNames(excludeColumnFiledNames);
        }
        if (CollUtil.isNotEmpty(dynamicHeads)) {
            sheetBuilder.registerWriteHandler(new DynamicHeadHandler(excelContext, dynamicHeads));
        }
        sheetBuilder.registerWriteHandler(new ExtendColumnHandler(excelContext));
        if (CollUtil.isNotEmpty(headStyles)) {
            sheetBuilder.registerWriteHandler(new HeadStyleHandler(excelContext, headStyles));
        } else {
            sheetBuilder.registerWriteHandler(new HeadStyleHandler(excelContext, headStyle));
        }
        sheetBuilder.registerWriteHandler(new ValidationInfoHandler(excelContext, validationInfos));
        sheetBuilder.registerWriteHandler(new ErrorInfoHandler(excelContext, removeOldErrorInfo, errorInfos));
        sheetBuilder.needHead(needHead);
        ZeusWriteSheet zeusWriteSheet = new ZeusWriteSheet(excelContext);
        BeanUtil.copyProperties(sheetBuilder.build(), zeusWriteSheet);
        return zeusWriteSheet;
    }

    public <T> void doWrite(Class<T> headClass, List<T> datas) {
        if (excelWriter == null) {
            throw new ExcelGenerateException("Must use 'ZeusExcel.write().sheet()' to call this method");
        }
        try {
            excelContext.setSheetData(datas);
            excelContext.setHeadClass(headClass);
            excelWriter.write(datas, build(headClass));
        } finally {
            if (excelWriter != null) {
                excelWriter.finish();
            }
            excelContext.clear();
        }
    }

    /**
     * @param headNames   表头
     * @param datas       没有指定module时要写入的数据，外层list下标表示行索引，内层list表示该行所有列的数据
     */
    public void doWrite(List<String> headNames, List<List<Object>> datas) {
        if (excelWriter == null) {
            throw new ExcelGenerateException("Must use 'ZeusExcel.write().sheet()' to call this method");
        }
        try {
            List<List<String>> heads = new ArrayList<>(headNames.size());
            headNames.forEach(head -> heads.add(new ArrayList<String>(1) {{add(head);}}));
            excelContext.setSheetData(datas);
            excelWriter.write(datas, build(heads));
        } finally {
            if (excelWriter != null) {
                excelWriter.finish();
            }
            excelContext.clear();
        }
    }

}
