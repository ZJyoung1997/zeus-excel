package com.jz.zeus.excel.test;

import com.alibaba.excel.context.AnalysisContext;
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
    protected void dataHandle(Map<Integer, DemoData> dataMap, AnalysisContext analysisContext) {
        System.out.println("加载数据：" + dataMap.values());
    }
}
