package com.jz.zeus.excel.write.handler;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.util.StrUtil;
import com.alibaba.excel.metadata.CellData;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder;
import com.alibaba.excel.write.property.ExcelWriteHeadProperty;
import com.jz.zeus.excel.DynamicHead;
import com.jz.zeus.excel.context.ExcelContext;
import org.apache.poi.ss.usermodel.Cell;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author:JZ
 * @date:2021/4/13
 */
public class DynamicHeadHandler extends AbstractCellWriteHandler {

    private boolean stopExecution;

    private ExcelContext excelContext;

    private List<DynamicHead> dynamicHeads;

    /**
     * 需要改变的表头
     */
    private Map<String, DynamicHead> needChangeHeadMap;


    public DynamicHeadHandler(ExcelContext excelContext, List<DynamicHead> dynamicHeads) {
        this.excelContext = excelContext;
        this.dynamicHeads = dynamicHeads;
    }

    @Override
    public void afterCellDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, List<CellData> cellDataList, Cell cell, Head head, Integer relativeRowIndex, Boolean isHead) {
        if (stopExecution || !Boolean.TRUE.equals(isHead)) {
            return;
        }
        String oldHeadName = cell.toString();
        DynamicHead dynamicHead = needChangeHeadMap.get(oldHeadName);
        if (dynamicHead != null) {
            cell.setCellValue(dynamicHead.buildFinalHeadName(oldHeadName));
        }
    }

    @Override
    public void beforeSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {
        if (CollUtil.isEmpty(dynamicHeads)) {
            stopExecution = true;
            return;
        }
        ExcelWriteHeadProperty excelWriteHeadProperty = writeSheetHolder.getExcelWriteHeadProperty();
        Map<Integer, Head> headMap = excelWriteHeadProperty.getHeadMap();
        if (CollUtil.isEmpty(headMap)) {
            stopExecution = true;
            return;
        }
        Map<String, DynamicHead> fieldHeadMap = new HashMap<>();
        Map<String, DynamicHead> headNameMap = new HashMap<>();
        Map<Integer, DynamicHead> indexHeadMap = new HashMap<>();
        dynamicHeads.forEach(dynamicHead -> {
            if (dynamicHead.getColumnIndex() != null) {
                indexHeadMap.put(dynamicHead.getColumnIndex(), dynamicHead);
            } else if (StrUtil.isNotBlank(dynamicHead.getFieldName())) {
                fieldHeadMap.put(dynamicHead.getFieldName(), dynamicHead);
            } else if (StrUtil.isNotBlank(dynamicHead.getHeadName())) {
                headNameMap.put(dynamicHead.getHeadName(), dynamicHead);
            }
        });

        if (needChangeHeadMap == null) {
            needChangeHeadMap = new HashMap<>();
        } else {
            needChangeHeadMap.clear();
        }
        Map<String, Integer> dynamicHeadIndexMap = new HashMap<>(dynamicHeads.size());
        for (Map.Entry<Integer, Head> entry : headMap.entrySet()) {
            Head head = entry.getValue();
            Integer columnIndex = entry.getKey();
            String fieldName = head.getFieldName();
            List<String> headNames = head.getHeadNameList();
            DynamicHead dynamicHead = indexHeadMap.get(columnIndex);
            if (dynamicHead != null && dynamicHead.getRowIndex() < headNames.size()) {
                needChangeHeadMap.put(headNames.get(dynamicHead.getRowIndex()), dynamicHead);
                dynamicHeadIndexMap.put(dynamicHead.buildFinalHeadName(headNames.get(dynamicHead.getRowIndex())), columnIndex);
            } else if ((dynamicHead = fieldHeadMap.get(fieldName)) != null && dynamicHead.getRowIndex() < headNames.size()) {
                needChangeHeadMap.put(headNames.get(dynamicHead.getRowIndex()), dynamicHead);
                dynamicHeadIndexMap.put(dynamicHead.buildFinalHeadName(headNames.get(dynamicHead.getRowIndex())), columnIndex);
            } else {
                for (int i = 0; i < headNames.size(); i++) {
                    dynamicHead = headNameMap.get(headNames.get(i));
                    if (dynamicHead != null && dynamicHead.getRowIndex() == i) {
                        needChangeHeadMap.put(headNames.get(i), dynamicHead);
                        dynamicHeadIndexMap.put(dynamicHead.buildFinalHeadName(headNames.get(i)), columnIndex);
                    }
                }
            }
        }
        excelContext.setDynamicHead(dynamicHeadIndexMap);
        if (needChangeHeadMap.size() == 0) {
            stopExecution = true;
        }
    }

}
