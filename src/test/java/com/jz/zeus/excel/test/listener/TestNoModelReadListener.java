package com.jz.zeus.excel.test.listener;

import com.alibaba.excel.context.AnalysisContext;
import com.jz.zeus.excel.read.listener.NoModelReadListener;

import java.util.Map;

/**
 * @author:JZ
 * @date:2021/4/4
 */
public class TestNoModelReadListener extends NoModelReadListener {

    @Override
    protected void dataHandle(Map<Integer, Map<Integer, String>> dataMap, Map<Integer, Map<String, String>> dynamicColumnDataMap,
                              AnalysisContext analysisContext) {
        dataMap.forEach((key, value) -> {
            System.out.println("加载class数据：" + value);
        });
        dynamicColumnDataMap.forEach((key, value) -> {
            value.forEach((k, v) -> {
                System.out.println(String.format("加载动态数据：%s -> %s", k, v));
            });
        });
    }

}
