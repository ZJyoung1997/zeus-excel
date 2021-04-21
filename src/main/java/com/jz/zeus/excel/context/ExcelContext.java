package com.jz.zeus.excel.context;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.map.MapUtil;
import com.jz.zeus.excel.DynamicHead;

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
            return CollUtil.newArrayList();
        }
        return result;
    }

    public static void setDynamicHead(List<DynamicHead> dynamicHeads) {
        EXCEL_CONTEXT.get().put(DYNAMIC_HEAD, dynamicHeads);
    }

    public static List<DynamicHead> getDynamicHead() {
        List<DynamicHead> dynamicHeads = (List<DynamicHead>) EXCEL_CONTEXT.get().get(DYNAMIC_HEAD);
        if (dynamicHeads == null) {
            return CollUtil.newArrayList();
        }
        return dynamicHeads;
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
