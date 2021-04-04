package com.jz.zeus.excel.read.listener;

import com.alibaba.excel.context.AnalysisContext;

import java.util.Map;

/**
 * 适用于没有定义 Model 的Excel读取
 * @author:JZ
 * @date:2021/4/4
 */
public abstract class NoModelReadListener extends AbstractExcelReadListener<Map<Integer, String>> {

    public NoModelReadListener() {}

    public NoModelReadListener(Boolean enabledAnnotationValidation) {
        super(enabledAnnotationValidation, null);
    }

    public NoModelReadListener(Integer batchHandleNum) {
        super(null, batchHandleNum);
    }

    public NoModelReadListener(Boolean enabledAnnotationValidation, Integer batchSaveNum) {
        super(enabledAnnotationValidation, batchSaveNum);
    }

    protected String getCellValue(Map<Integer, String> rowValueMap, String headName) {
        return rowValueMap.get(getHeadIndex(headName));
    }

}
