package com.jz.zeus.excel.read.listener;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.metadata.CellData;
import com.alibaba.excel.util.ConverterUtils;

import java.util.Map;

/**
 * @Author JZ
 * @Date 2021/3/31 16:48
 */
public abstract class AbstractExcelReadListener<T> extends ExcelReadListener<T> {

    public AbstractExcelReadListener() {}

    public AbstractExcelReadListener(Boolean enabledAnnotationValidation) {
        super(enabledAnnotationValidation, null);
    }

    public AbstractExcelReadListener(Integer batchHandleNum) {
        super(null, batchHandleNum);
    }

    public AbstractExcelReadListener(Boolean enabledAnnotationValidation, Integer batchSaveNum) {
        super(enabledAnnotationValidation, batchSaveNum);
    }

    @Override
    protected void doAfterAllDataHandle(AnalysisContext analysisContext) {}

    @Override
    protected void verify(Map<Integer, T> dataMap, AnalysisContext analysisContext) {}

    @Override
    protected void headCheck(Map<Integer, CellData> headMap, AnalysisContext context) {
        headCheck(ConverterUtils.convertToStringMap(headMap, context));
    }

    protected void headCheck(Map<Integer, String> headMap) {}

}
