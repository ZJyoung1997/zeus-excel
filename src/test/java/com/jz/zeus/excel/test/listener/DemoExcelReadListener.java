package com.jz.zeus.excel.test.listener;

import cn.hutool.core.lang.Console;
import com.alibaba.excel.context.AnalysisContext;
import com.jz.zeus.excel.read.listener.AbstractExcelReadListener;
import com.jz.zeus.excel.test.data.DemoData;

import java.util.Map;

/**
 * @Author JZ
 * @Date 2021/3/22 15:23
 */
public class DemoExcelReadListener extends AbstractExcelReadListener<DemoData> {

    public DemoExcelReadListener() {}

    public DemoExcelReadListener(Boolean lastHandleData) {
        super(lastHandleData);
    }

    public DemoExcelReadListener(Integer batchSaveNum) {
        super(batchSaveNum);
    }

    @Override
    protected void verify(Map<Integer, DemoData> dataMap, AnalysisContext analysisContext) {
        dataMap.forEach((rowIndex, data) -> {
            Console.log();
        });
    }

    @Override
    protected void dataHandle(Map<Integer, DemoData> dataMap, AnalysisContext analysisContext) {
//        dataMap.forEach((key, value) -> {
//            System.out.println("加载class数据：" + value);
//        });
    }

}
