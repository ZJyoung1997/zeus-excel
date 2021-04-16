package com.jz.zeus.excel.read.listener;

import cn.hutool.core.collection.CollUtil;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.metadata.CellData;

import java.util.HashMap;
import java.util.Map;

/**
 * 适用于没有定义 Model 的Excel读取
 * @author:JZ
 * @date:2021/4/4
 */
public abstract class NoModelReadListener extends AbstractExcelReadListener<Map<Integer, String>> {

    /**
     * key 表头、value 表头索引
     */
    private Map<String, Integer> headNameIndexMap = new HashMap<>();

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


    @Override
    protected void headCheck(Map<Integer, CellData> headMap, AnalysisContext context) {
        super.headCheck(headMap, context);
        if (CollUtil.isEmpty(headMap)) {
            return;
        }
        headMap.forEach((columnIndex, cell) -> {
            headNameIndexMap.put(cell.toString(), columnIndex);
        });
    }

    protected String getCellValue(Map<Integer, String> rowValueMap, String headName) {
        return rowValueMap.get(headNameIndexMap.get(headName));
    }

}
