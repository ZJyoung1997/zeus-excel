package com.jz.zeus.excel.test.listener;

import cn.hutool.core.lang.Console;
import com.alibaba.excel.context.AnalysisContext;
import com.jz.zeus.excel.read.listener.ExcelReadListener;
import com.jz.zeus.excel.test.data.DemoData;

import java.util.Map;

public class DemoReadListener extends ExcelReadListener<DemoData> {

    @Override
    protected void headHandle(Map<Integer, String> headMap, AnalysisContext analysisContext) {
        Console.log("sheet表头：{}", headMap.values());
    }

    @Override
    protected void dataHandle(Map<Integer, DemoData> dataMap, AnalysisContext analysisContext) {
        dataMap.forEach((key, value) -> {
            Console.log("第{}行数据：{}", key + 1, value);
        });
    }

    @Override
    protected void doAfterAllDataHandle(AnalysisContext analysisContext) {
        Console.log("Excel handle done");
    }

}
