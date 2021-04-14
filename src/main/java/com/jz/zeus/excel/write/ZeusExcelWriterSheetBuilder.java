package com.jz.zeus.excel.write;

import cn.hutool.core.collection.CollUtil;
import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.write.builder.ExcelWriterSheetBuilder;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.jz.zeus.excel.CellErrorInfo;
import com.jz.zeus.excel.DynamicHead;
import com.jz.zeus.excel.ValidationInfo;
import com.jz.zeus.excel.write.handler.DynamicHeadHandler;
import com.jz.zeus.excel.write.handler.ExtendColumnHandler;

import java.util.List;

/**
 * @author:JZ
 * @date:2021/4/14
 */
public class ZeusExcelWriterSheetBuilder {

    private Class<?> headClass;

    private Integer sheetIndex;

    private String sheetName;

    private List<String> extendHead;

    private List<DynamicHead> dynamicHeads;

    private List<ValidationInfo> validationInfos;

    private List<CellErrorInfo> errorInfos;

    public ZeusExcelWriterSheetBuilder head(Class<?> headClass) {
        this.headClass = headClass;
        return this;
    }

    public ZeusExcelWriterSheetBuilder sheet() {
        this.sheetIndex = 0;
        return this;
    }

    public ZeusExcelWriterSheetBuilder sheet(Integer sheetIndex) {
        this.sheetIndex = sheetIndex;
        return this;
    }

    public ZeusExcelWriterSheetBuilder sheet(String sheetName) {
        this.sheetName = sheetName;
        return this;
    }

    public ZeusExcelWriterSheetBuilder sheet(Integer sheetIndex, String sheetName) {
        this.sheetIndex = sheetIndex;
        this.sheetName = sheetName;
        return this;
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

    public WriteSheet build() {
        ExcelWriterSheetBuilder sheetBuilder = EasyExcel.writerSheet(sheetIndex, sheetName);
        if (CollUtil.isNotEmpty(dynamicHeads)) {
            sheetBuilder.registerWriteHandler(new DynamicHeadHandler(dynamicHeads));
        }
        if (CollUtil.isNotEmpty(extendHead)) {
            sheetBuilder.registerWriteHandler(new ExtendColumnHandler<>(headClass))
        }
    }

}
