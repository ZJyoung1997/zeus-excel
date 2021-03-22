package com.jz.zeus.excel.test;

import com.alibaba.excel.context.AnalysisContext;
import com.jz.zeus.excel.read.listener.AbstractExcelReadListener;

/**
 * @Author JZ
 * @Date 2021/3/22 15:23
 */
public class DemoExcelReadListener extends AbstractExcelReadListener<DemoData> {

    public DemoExcelReadListener(boolean lastSave) {
        this.lastSave = lastSave;
    }

    public DemoExcelReadListener(int batchSaveNum) {
        this.batchSaveNum = batchSaveNum;
    }

    public DemoExcelReadListener(boolean lastSave, int batchSaveNum) {
        this.lastSave = lastSave;
        this.batchSaveNum = batchSaveNum;
    }

    @Override
    protected void save(AnalysisContext analysisContext) {
        System.out.println("加载数据：" + dataList);
    }

}
