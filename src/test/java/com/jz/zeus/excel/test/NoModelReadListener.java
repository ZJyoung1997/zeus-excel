package com.jz.zeus.excel.test;

import com.alibaba.excel.context.AnalysisContext;
import com.jz.zeus.excel.read.listener.AbstractExcelReadListener;

import java.util.Map;

/**
 * @Author JZ
 * @Date 2021/4/2 11:37
 */
public class NoModelReadListener extends AbstractExcelReadListener<Map<Integer, String>> {

    @Override
    protected void dataHandle(Map<Integer, Map<Integer, String>> dataMap, AnalysisContext analysisContext) {

    }

}
