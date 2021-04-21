package com.jz.zeus.excel.write.handler;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.util.StrUtil;
import com.alibaba.excel.metadata.CellData;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.write.handler.AbstractCellWriteHandler;
import com.alibaba.excel.write.handler.AbstractSheetWriteHandler;
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
import java.util.Objects;

/**
 * @author:JZ
 * @date:2021/4/13
 */
public class DynamicHeadHandler extends AbstractCellWriteHandler {

    private boolean isEmpty;

    private Map<String, DynamicHead> fieldHeadMap;

    private Map<String, DynamicHead> headNameMap;

    private Map<Integer, DynamicHead> indexHeadMap;

    public DynamicHeadHandler(List<DynamicHead> dynamicHeads) {
        isEmpty = CollUtil.isEmpty(dynamicHeads);
        if (isEmpty) {
            return;
        }
        ExcelContext.setDynamicHead(dynamicHeads);

        fieldHeadMap = new HashMap<>();
        headNameMap = new HashMap<>();
        indexHeadMap = new HashMap<>();
        dynamicHeads.forEach(dynamicHead -> {
            if (dynamicHead.getColumnIndex() != null) {
                indexHeadMap.put(dynamicHead.getColumnIndex(), dynamicHead);
            } else if (StrUtil.isNotBlank(dynamicHead.getFieldName())) {
                fieldHeadMap.put(dynamicHead.getFieldName(), dynamicHead);
            }
        });
    }

    @Override
    public void afterCellDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, List<CellData> cellDataList, Cell cell, Head head, Integer relativeRowIndex, Boolean isHead) {
        if (isEmpty || !Boolean.TRUE.equals(isHead)) {
            return;
        }
        Integer rowIndex = cell.getRowIndex();
        Integer columnIndex = head.getColumnIndex();
        String fieldName = head.getFieldName();
        List<String> headNames = head.getHeadNameList();

        DynamicHead dynamicHead = indexHeadMap.get(columnIndex);
        if (columnIndex != null && dynamicHead != null && Objects.equals(dynamicHead.getRowIndex(), rowIndex)) {
            cell.setCellValue(dynamicHead.getFinalHeadName(cell.toString()));
        } else if (StrUtil.isNotBlank(fieldName) && (dynamicHead = fieldHeadMap.get(fieldName)) != null
                && Objects.equals(dynamicHead.getRowIndex(), rowIndex)) {
            cell.setCellValue(dynamicHead.getFinalHeadName(cell.toString()));
        } else if (CollUtil.isNotEmpty(headNames)) {
            for (int i = 0; i < headNames.size(); i++) {
                dynamicHead = headNameMap.get(headNames.get(i));
                if (i != rowIndex || dynamicHead == null || !Objects.equals(dynamicHead.getRowIndex(), rowIndex)) {
                    continue;
                }
                cell.setCellValue(dynamicHead.getFinalHeadName(cell.toString()));
            }
        }

    }
}
