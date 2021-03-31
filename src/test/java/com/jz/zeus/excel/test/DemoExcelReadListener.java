package com.jz.zeus.excel.test;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.metadata.CellData;
import com.jz.zeus.excel.read.listener.AbstractExcelReadListener;

import java.util.Map;

/**
 * @Author JZ
 * @Date 2021/3/22 15:23
 */
public class DemoExcelReadListener extends AbstractExcelReadListener<DemoData> {

    public DemoExcelReadListener(Boolean lastHandleData) {
        super(lastHandleData);
    }

    public DemoExcelReadListener(Integer batchSaveNum) {
        super(batchSaveNum);
    }

    @Override
    protected void dataHandle(AnalysisContext analysisContext, Integer currentRowIndex) {
        System.out.println("加载数据：" + dataList);
    }

    @Override
    protected void doAfterAllDataHandle(AnalysisContext analysisContext) {

    }

    @Override
    protected void verify(AnalysisContext analysisContext, Integer currentRowIndex) {

    }

    @Override
    protected void headCheck(Map<Integer, CellData> headMap, AnalysisContext context) {

    }

}
