package com.jz.zeus.excel.write.handler;

import cn.hutool.core.collection.CollUtil;
import com.alibaba.excel.annotation.write.style.*;
import com.alibaba.excel.enums.HeadKindEnum;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.metadata.property.ColumnWidthProperty;
import com.alibaba.excel.metadata.property.FontProperty;
import com.alibaba.excel.metadata.property.LoopMergeProperty;
import com.alibaba.excel.metadata.property.StyleProperty;
import com.alibaba.excel.write.handler.AbstractCellWriteHandler;
import com.alibaba.excel.write.handler.AbstractRowWriteHandler;
import com.alibaba.excel.write.handler.SheetWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder;
import com.alibaba.excel.write.property.ExcelWriteHeadProperty;
import com.jz.zeus.excel.annotation.DynamicColumn;
import com.jz.zeus.excel.util.UnsafeFieldAccessor;
import lombok.SneakyThrows;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.lang.reflect.Field;
import java.util.*;

/**
 * @Author JZ
 * @Date 2021/4/12 15:23
 */
public class ExtendColumnHandler<T> extends AbstractRowWriteHandler implements SheetWriteHandler {

    private Field dynamicColumnField;

    private Integer lastColumnIndex;

    private boolean isClassHead;

    /**
     * 表头行数
     */
    private int headRowNum;

    private Map<Integer, T> dataMap;

    private UnsafeFieldAccessor fieldAccessor;

    public ExtendColumnHandler(List<T> dataList) {
        dataMap = new HashMap<>(dataList == null ? 0 : dataList.size());
        if (CollUtil.isNotEmpty(dataList)) {
            for (int i = 0; i < dataList.size(); i++) {
                dataMap.put(i, dataList.get(i));
            }
        }
    }

    @Override
    public void afterRowDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, Row row, Integer relativeRowIndex, Boolean isHead) {
        if (Boolean.TRUE.equals(isHead) || !isClassHead || dynamicColumnField == null) {
            return;
        }
        T rawData = dataMap.get(row.getRowNum()-headRowNum);
        if (rawData == null) {
            return;
        }
        Map<String, String> dynamicData = (Map<String, String>) fieldAccessor.getObject(rawData);
        if (CollUtil.isEmpty(dynamicData)) {
            return;
        }
        int index = lastColumnIndex;
        for (String value : dynamicData.values()) {
            Cell cell = row.createCell(index++);
            cell.setCellValue(value);
        }
    }

    @Override
    public void afterSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {
        ExcelWriteHeadProperty excelWriteHeadProperty = writeSheetHolder.getExcelWriteHeadProperty();
        lastColumnIndex = excelWriteHeadProperty.getHeadMap().size();
        isClassHead = HeadKindEnum.CLASS.equals(excelWriteHeadProperty.getHeadKind());
        headRowNum = excelWriteHeadProperty.getHeadRowNumber();

        Map<String, Field> ignoreFieldMap = excelWriteHeadProperty.getIgnoreMap();
        if (CollUtil.isEmpty(ignoreFieldMap)) {
            return;
        }
        for (Field field : ignoreFieldMap.values()) {
            if (field.getAnnotation(DynamicColumn.class) != null) {
                dynamicColumnField = field;
                dynamicColumnField.setAccessible(true);
                fieldAccessor = new UnsafeFieldAccessor(dynamicColumnField);
                break;
            }
        }
        if (dynamicColumnField == null) {
            return;
        }
        addHead(excelWriteHeadProperty);
    }

    private void addHead(ExcelWriteHeadProperty excelWriteHeadProperty) {
        Map<Integer, Head> headMap = excelWriteHeadProperty.getHeadMap();
        T rawData = dataMap.get(0);
        if (rawData == null) {
            return;
        }
        Map<String, String> dynamicData = (Map<String, String>) fieldAccessor.getObject(rawData);
        if (CollUtil.isEmpty(dynamicData)) {
            return;
        }
        FontProperty headFontProperty = FontProperty.build(dynamicColumnField.getAnnotation(HeadFontStyle.class));
        StyleProperty headStyleProperty = StyleProperty.build(dynamicColumnField.getAnnotation(HeadStyle.class));
        FontProperty contentFontProperty = FontProperty.build(dynamicColumnField.getAnnotation(ContentFontStyle.class));
        StyleProperty contentStyleProperty = StyleProperty.build(dynamicColumnField.getAnnotation(ContentStyle.class));
        ColumnWidthProperty columnWidthProperty = ColumnWidthProperty.build(dynamicColumnField.getAnnotation(ColumnWidth.class));
        LoopMergeProperty loopMergeProperty = LoopMergeProperty.build(dynamicColumnField.getAnnotation(ContentLoopMerge.class));
        int index = lastColumnIndex;
        for (String headName : dynamicData.keySet()) {
            List<String> headNames = new ArrayList<>(headRowNum);
            for (int i = 0; i < headRowNum; i++) {
                headNames.add(headName);
            }
            Head head = new Head(index, null, headNames, false, true);
            head.setHeadFontProperty(headFontProperty);
            head.setHeadStyleProperty(headStyleProperty);
            head.setContentFontProperty(contentFontProperty);
            head.setContentStyleProperty(contentStyleProperty);
            head.setColumnWidthProperty(columnWidthProperty);
            head.setLoopMergeProperty(loopMergeProperty);
            headMap.put(index, head);
            index++;
        }
    }

    @Override
    public void beforeSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {}

}
