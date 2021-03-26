package com.jz.zeus.excel.write.handler;

import cn.hutool.core.collection.CollUtil;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.write.handler.AbstractSheetWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;

import java.util.HashMap;
import java.util.Map;

/**
 * @Author JZ
 * @Date 2021/3/26 17:11
 */
public class AbstractZeusSheetWriteHandler extends AbstractSheetWriteHandler {

    /**
     * key 表头、value 表头对应列索引
     */
    private Map<String, Integer> headNameIndexMap;

    protected void initHeadNameIndexMap(WriteSheetHolder writeSheetHolder) {
        Map<Integer, Head> headMap = writeSheetHolder.getExcelWriteHeadProperty().getHeadMap();
        headNameIndexMap = new HashMap<>();
        headMap.values().forEach(head -> {
            Integer columnIndex = head.getColumnIndex();
            head.getHeadNameList().forEach(headName -> {
                headNameIndexMap.put(headName, columnIndex);
            });
        });
    }

    protected Integer getHeadColumnIndex(String headName) {
        return CollUtil.isEmpty(headNameIndexMap) ? null : headNameIndexMap.get(headName);
    }

}
