package com.jz.zeus.excel.read.listener;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.metadata.CellData;

import java.util.Map;

/**
 * @Author JZ
 * @Date 2021/3/31 16:48
 */
public abstract class AbstractExcelReadListener<T> extends ExcelReadListener<T> {

    public AbstractExcelReadListener() {}

    public AbstractExcelReadListener(Boolean lastHandleData) {
        super(lastHandleData, null, null);
    }

    public AbstractExcelReadListener(Integer batchSaveNum) {
        super(null, null, batchSaveNum);
    }

    public AbstractExcelReadListener(Boolean lastHandleData, Boolean enabledAnnotationValidation, Integer batchSaveNum) {
        super(lastHandleData, enabledAnnotationValidation, batchSaveNum);
    }

    @Override
    protected void doAfterAllDataHandle(AnalysisContext analysisContext) {}

    @Override
    protected void verify(AnalysisContext analysisContext, Integer currentRowIndex) {}

    @Override
    protected void headCheck(Map<Integer, CellData> headMap, AnalysisContext context) {}

}
