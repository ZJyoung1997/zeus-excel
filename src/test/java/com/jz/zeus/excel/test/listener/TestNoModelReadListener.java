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
    protected void dataHandle(Map<Integer, Map<Integer, String>> dataMap, AnalysisContext analysisContext) {
        dataMap.values().forEach(data -> {
            System.out.println("读取到数据：" + data.values());
        });
    }

}
