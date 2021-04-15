package com.jz.zeus.excel.write.builder;

import cn.hutool.core.collection.CollUtil;
import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.exception.ExcelGenerateException;
import com.alibaba.excel.write.builder.ExcelWriterSheetBuilder;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.jz.zeus.excel.CellErrorInfo;
import com.jz.zeus.excel.DynamicHead;
import com.jz.zeus.excel.ValidationInfo;
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

    private ExcelWriter excelWriter;

    private Integer sheetIndex;

    private String sheetName;

    private List<String> extendHead;

    private List<DynamicHead> dynamicHeads;

    private List<ValidationInfo> validationInfos;

    private List<CellErrorInfo> errorInfos;

    private CellStyleProperty headStyle;

    private List<CellStyleProperty> singleRowHeadStyles;

    private List<List<CellStyleProperty>> multiRowHeadStyles;

    private List<String> excludeColumnFiledNames;


    public ZeusExcelWriterSheetBuilder(Integer sheetIndex, String sheetName) {
        this.sheetIndex = sheetIndex == null ? SHEET_INDEX_DEFAULT : sheetIndex;
        this.sheetName = sheetName;
    }

    public ZeusExcelWriterSheetBuilder(ExcelWriter excelWriter, Integer sheetIndex, String sheetName) {
        this.excelWriter = excelWriter;
        this.sheetIndex = sheetIndex == null ? SHEET_INDEX_DEFAULT : sheetIndex;
        this.sheetName = sheetName;
    }

    public ZeusExcelWriterSheetBuilder extendHead(List<String> extendHead) {
        this.extendHead = extendHead;
        return this;
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

    public ZeusExcelWriterSheetBuilder singleRowHeadStyles(List<CellStyleProperty> singleRowHeadStyles) {
        this.singleRowHeadStyles = singleRowHeadStyles;
        return this;
    }

    public ZeusExcelWriterSheetBuilder multiRowHeadStyles(List<List<CellStyleProperty>> multiRowHeadStyles) {
        this.multiRowHeadStyles = multiRowHeadStyles;
        return this;
    }

    public ZeusExcelWriterSheetBuilder excludeColumnFiledNames(List<String> excludeColumnFiledNames) {
        this.excludeColumnFiledNames = excludeColumnFiledNames;
        return this;
    }

    public WriteSheet build(List<List<String>> headNames) {
        ExcelWriterSheetBuilder sheetBuilder = EasyExcel.writerSheet(sheetIndex, sheetName);
        sheetBuilder.head(headNames);
        if (CollUtil.isNotEmpty(validationInfos)) {
            sheetBuilder.registerWriteHandler(new ValidationInfoHandler(validationInfos));
        }
        if (CollUtil.isNotEmpty(errorInfos)) {
            sheetBuilder.registerWriteHandler(new ErrorInfoHandler(errorInfos));
        }
        return sheetBuilder.build();
    }

    public WriteSheet build(Class headClass, List datas) {
        ExcelWriterSheetBuilder sheetBuilder = EasyExcel.writerSheet(sheetIndex, sheetName);
        sheetBuilder.head(headClass);
        if (CollUtil.isNotEmpty(excludeColumnFiledNames)) {
            sheetBuilder.excludeColumnFiledNames(excludeColumnFiledNames);
        }
        if (CollUtil.isNotEmpty(dynamicHeads)) {
            sheetBuilder.registerWriteHandler(new DynamicHeadHandler(dynamicHeads));
        }
        if (CollUtil.isNotEmpty(extendHead)) {
            sheetBuilder.registerWriteHandler(new ExtendColumnHandler(datas, extendHead));
        }
        if (CollUtil.isNotEmpty(multiRowHeadStyles)) {
            sheetBuilder.registerWriteHandler(new HeadStyleHandler().setMultiRowHeadCellStyles(multiRowHeadStyles));
        } else if (CollUtil.isNotEmpty(singleRowHeadStyles)) {
            sheetBuilder.registerWriteHandler(new HeadStyleHandler(singleRowHeadStyles));
        } else {
            sheetBuilder.registerWriteHandler(new HeadStyleHandler(headStyle));
        }
        if (CollUtil.isNotEmpty(validationInfos)) {
            sheetBuilder.registerWriteHandler(new ValidationInfoHandler(validationInfos));
        }
        if (CollUtil.isNotEmpty(errorInfos)) {
            sheetBuilder.registerWriteHandler(new ErrorInfoHandler(errorInfos));
        }
        return sheetBuilder.build();
    }

    public <T> void doWrite(Class<T> headClass, List<? extends T> datas) {
        if (excelWriter == null) {
            throw new ExcelGenerateException("Must use 'ZeusExcel.write().sheet()' to call this method");
        }
        try {
            excelWriter.write(datas, build(headClass, datas));
        } finally {
            if (excelWriter != null) {
                excelWriter.finish();
            }
        }
    }

    /**
     * @param headNames   表头
     * @param dataList    要写入的数据，外层list下标表示行索引，内层list表示该行所有列的数据
     */
    public void doWrite(List<String> headNames, List<List<Object>> dataList) {
        if (excelWriter == null) {
            throw new ExcelGenerateException("Must use 'ZeusExcel.write().sheet()' to call this method");
        }
        try {
            List<List<String>> heads = new ArrayList<>(headNames.size());
            headNames.forEach(head -> heads.add(new ArrayList<String>(1) {{add(head);}}));
            excelWriter.write(dataList, build(heads));
        } finally {
            if (excelWriter != null) {
                excelWriter.finish();
            }
        }

    }

}
