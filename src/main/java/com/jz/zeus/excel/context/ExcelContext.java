package com.jz.zeus.excel.context;

import cn.hutool.core.map.MapUtil;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @Author JZ
 * @Date 2021/4/21 15:47
 */
public class ExcelContext {

    private static final String HEAD_CLASS = "head_class";

    private static final String EXTEND_HEAD = "extend_head";

    private static final String DYNAMIC_HEAD = "dynamic_head";

    private static ThreadLocal<Map<String, Object>> EXCEL_CONTEXT = ThreadLocal.withInitial(() -> new HashMap<>());

    public static void setExtendHead(List<String> extendHeads) {
        EXCEL_CONTEXT.get().put(EXTEND_HEAD, extendHeads);
    }

    public static List<String> getExtendHead() {
        List<String> result = (List<String>) EXCEL_CONTEXT.get().get(EXTEND_HEAD);
        if (result == null) {
            return new ArrayList<>(0);
        }
        return result;
    }

    public static void setDynamicHead(Map<String, Integer> dynamicHeadIndexMap) {
        EXCEL_CONTEXT.get().put(DYNAMIC_HEAD, dynamicHeadIndexMap);
    }

    public static Map<String, Integer> getDynamicHead() {
        Map<String, Integer> dynamicHeadIndexMap = (Map<String, Integer>) EXCEL_CONTEXT.get().get(DYNAMIC_HEAD);
        if (dynamicHeadIndexMap == null) {
            return MapUtil.newHashMap();
        }
        return dynamicHeadIndexMap;
    }

    public static void setHeadClass(Class<?> headClass) {
        EXCEL_CONTEXT.get().put(HEAD_CLASS, headClass);
    }

    public static Class getHeadClass() {
        return (Class) EXCEL_CONTEXT.get().get(HEAD_CLASS);
    }

    public static void clear() {
        EXCEL_CONTEXT.get().clear();
    }

}
