package com.jz.zeus.excel.test.listener;

import com.alibaba.excel.context.AnalysisContext;
import com.jz.zeus.excel.read.listener.AbstractExcelReadListener;
import com.jz.zeus.excel.test.data.DemoData;

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
    protected void verify(Map<Integer, DemoData> dataMap, AnalysisContext analysisContext) {
        dataMap.forEach((rowIndex, data) -> {
            addErrorInfoByHead(rowIndex, "自定义1", "格式错误");
        });
    }

    @Override
    protected void dataHandle(Map<Integer, DemoData> dataMap, AnalysisContext analysisContext) {
        dataMap.forEach((key, value) -> {
            System.out.println("加载class数据：" + value);
        });
    }

}
