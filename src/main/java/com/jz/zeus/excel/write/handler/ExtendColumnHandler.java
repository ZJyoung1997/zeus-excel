package com.jz.zeus.excel.write.handler;

import cn.hutool.core.collection.CollUtil;
import com.alibaba.excel.enums.HeadKindEnum;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.write.handler.AbstractRowWriteHandler;
import com.alibaba.excel.write.handler.SheetWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder;
import com.alibaba.excel.write.property.ExcelWriteHeadProperty;
import com.jz.zeus.excel.FieldInfo;
import com.jz.zeus.excel.util.ClassUtils;
import com.jz.zeus.excel.util.UnsafeFieldAccessor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * @Author JZ
 * @Date 2021/4/12 15:23
 */
public class ExtendColumnHandler<T> extends AbstractRowWriteHandler implements SheetWriteHandler {

    private Field dynamicColumnField;

    private Integer lastColumnIndex;

    private boolean isClassHead;

    private Map<String, Integer> dynamicHeadIndexMap = new HashMap<>();

    /**
     * 表头行数
     */
    private int headRowNum;

    private Map<Integer, T> dataMap;

    /**
     * 动态表头
     */
    private List<String> dynamicHead;

    private UnsafeFieldAccessor fieldAccessor;

    public ExtendColumnHandler(List<T> dataList, List<String> dynamicHead) {
        dataMap = new HashMap<>(dataList == null ? 0 : dataList.size());
        if (CollUtil.isNotEmpty(dataList)) {
            for (int i = 0; i < dataList.size(); i++) {
                dataMap.put(i, dataList.get(i));
            }
        }
        if (CollUtil.isNotEmpty(dynamicHead)) {
            this.dynamicHead = new ArrayList<>(dynamicHead);
        }
    }

    @Override
    public void afterRowDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, Row row, Integer relativeRowIndex, Boolean isHead) {
        if (CollUtil.isEmpty(dynamicHead) || Boolean.TRUE.equals(isHead)
                || !isClassHead || dynamicColumnField == null) {
            return;
        }
        T rawData = dataMap.get(row.getRowNum()-headRowNum);
        if (rawData == null) {
            return;
        }
        Map<String, String> dynamicData = getDynamicData(rawData);
        if (CollUtil.isEmpty(dynamicData)) {
            return;
        }
        dynamicData.forEach((head, data) -> {
            Integer columnIndex = dynamicHeadIndexMap.get(head);
            if (columnIndex != null) {
                Cell cell = row.createCell(columnIndex);
                cell.setCellValue(data);
            }
        });
    }

    @Override
    public void afterSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {
        if (CollUtil.isEmpty(dynamicHead) && CollUtil.isEmpty(dataMap)) {
            return;
        }
        ExcelWriteHeadProperty excelWriteHeadProperty = writeSheetHolder.getExcelWriteHeadProperty();
        lastColumnIndex = excelWriteHeadProperty.getHeadMap().size();
        isClassHead = HeadKindEnum.CLASS.equals(excelWriteHeadProperty.getHeadKind());
        headRowNum = excelWriteHeadProperty.getHeadRowNumber();

        List<FieldInfo> fieldInfos = ClassUtils.getClassFieldInfo(writeSheetHolder.getClazz());
        fieldInfos.stream().filter(FieldInfo::isDynamicColumn)
                .findFirst().ifPresent(fieldInfo -> {
            dynamicColumnField = fieldInfo.getField();
            dynamicColumnField.setAccessible(true);
            fieldAccessor = new UnsafeFieldAccessor(dynamicColumnField);
            addHead(excelWriteHeadProperty, fieldInfo);
        });
    }

    private void addHead(ExcelWriteHeadProperty excelWriteHeadProperty, FieldInfo fieldInfo) {
        Map<Integer, Head> headMap = excelWriteHeadProperty.getHeadMap();
        if (CollUtil.isEmpty(dynamicHead)) {
            dynamicHead = new ArrayList<>();
        }
        if (CollUtil.isNotEmpty(dataMap)) {
            T rawData = dataMap.get(0);
            if (rawData != null) {
                Map<String, String> dynamicData = getDynamicData(rawData);
                if (CollUtil.isNotEmpty(dynamicData)) {
                    dynamicHead.addAll(dynamicData.keySet());
                    dynamicHead = dynamicHead.stream().distinct()
                            .collect(Collectors.toList());
                }
            }
        }
        if (dynamicHead.size() == 0) {
            return;
        }
        int index = lastColumnIndex;
        for (String headName : dynamicHead) {
            List<String> headNames = new ArrayList<>(headRowNum);
            for (int i = 0; i < headRowNum; i++) {
                headNames.add(headName);
            }
            dynamicHeadIndexMap.put(headName, index);
            Head head = new Head(index, null, headNames, false, true);
            head.setHeadFontProperty(fieldInfo.getHeadFontProperty());
            head.setHeadStyleProperty(fieldInfo.getHeadStyleProperty());
            head.setContentFontProperty(fieldInfo.getContentFontProperty());
            head.setContentStyleProperty(fieldInfo.getContentStyleProperty());
            head.setColumnWidthProperty(fieldInfo.getColumnWidthProperty());
            head.setLoopMergeProperty(fieldInfo.getLoopMergeProperty());
            headMap.put(index, head);
            index++;
        }
    }

    private Map<String, String> getDynamicData(T rawData) {
        return (Map<String, String>) fieldAccessor.getObject(rawData);
    }

    @Override
    public void beforeSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {}

}
