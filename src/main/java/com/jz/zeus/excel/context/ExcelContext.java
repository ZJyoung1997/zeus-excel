package com.jz.zeus.excel.context;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.collection.ListUtil;
import cn.hutool.core.map.MapUtil;

import java.util.Collection;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @Author JZ
 * @Date 2021/4/21 15:47
 */
public class ExcelContext {

    private final String HEAD_CLASS = "head_class";

    private final String EXTEND_HEAD = "extend_head";

    private final String DYNAMIC_HEAD = "dynamic_head";

    private final String SHEET_DATA = "sheet_data";

    private Map<String, Object> EXCEL_CONTEXT = new HashMap<>(4);

    public void setExtendHead(Map<String, Integer> extendHeadMap) {
        EXCEL_CONTEXT.put(EXTEND_HEAD, extendHeadMap);
    }

    public Map<String, Integer> getExtendHead() {
        Map<String, Integer> result = (Map<String, Integer>) EXCEL_CONTEXT.get(EXTEND_HEAD);
        if (result == null) {
            return MapUtil.newHashMap(0);
        }
        return result;
    }

    public void setDynamicHead(Map<String, Integer> dynamicHeadIndexMap) {
        EXCEL_CONTEXT.put(DYNAMIC_HEAD, dynamicHeadIndexMap);
    }

    public Map<String, Integer> getDynamicHead() {
        Map<String, Integer> dynamicHeadIndexMap = (Map<String, Integer>) EXCEL_CONTEXT.get(DYNAMIC_HEAD);
        if (dynamicHeadIndexMap == null) {
            return MapUtil.newHashMap();
        }
        return dynamicHeadIndexMap;
    }

    public void setHeadClass(Class<?> headClass) {
        EXCEL_CONTEXT.put(HEAD_CLASS, headClass);
    }

    public Class getHeadClass() {
        return (Class) EXCEL_CONTEXT.get(HEAD_CLASS);
    }

    public void setSheetData(List data) {
        EXCEL_CONTEXT.put(SHEET_DATA, ListUtil.toList(data));
    }

    public void addSheetData(List data) {
        List rawData = (List) EXCEL_CONTEXT.get(SHEET_DATA);
        if (rawData == null) {
            EXCEL_CONTEXT.put(SHEET_DATA, ListUtil.toList(data));
        } else {
            rawData.addAll(data);
        }
    }

    public List getSheetData() {
        List data = (List) EXCEL_CONTEXT.get(SHEET_DATA);
        if (data == null) {
            return ListUtil.list(false);
        }
        return data;
    }

    public void clear() {
        for (Map.Entry<String, Object> entry : EXCEL_CONTEXT.entrySet()) {
            Object cache = entry.getValue();
            if (cache instanceof Collection) {
                CollUtil.clear((Collection) cache);
            } else if (cache instanceof Map) {
                MapUtil.clear((Map) cache);
            } else {
                EXCEL_CONTEXT.put(entry.getKey(), null);
            }
        }
        EXCEL_CONTEXT.clear();
    }

}
