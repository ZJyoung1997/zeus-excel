package com.jz.zeus.excel.write.handler;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.collection.ListUtil;
import com.alibaba.excel.enums.HeadKindEnum;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.write.handler.AbstractRowWriteHandler;
import com.alibaba.excel.write.handler.SheetWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder;
import com.alibaba.excel.write.property.ExcelWriteHeadProperty;
import com.jz.zeus.excel.FieldInfo;
import com.jz.zeus.excel.context.ExcelContext;
import com.jz.zeus.excel.util.ClassUtils;
import com.jz.zeus.excel.util.UnsafeFieldAccessor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * 当存在动态表头时，该 handler应该最先被注册，已保证将动态表头添加到配置中
 * @Author JZ
 * @Date 2021/4/12 15:23
 */
public class ExtendColumnHandler extends AbstractRowWriteHandler implements SheetWriteHandler {

    private ExcelContext excelContext;

    private Field extendColumnField;

    private boolean isNotClassHead;

    private Map<String, Integer> extendHeadIndexMap = new HashMap<>();

    /**
     * 表头行数
     */
    private int headRowNum;

    private Map<Integer, Object> dataMap;

    /**
     * 动态表头
     */
    private List<String> extendHead;

    private UnsafeFieldAccessor fieldAccessor;

    public ExtendColumnHandler(ExcelContext excelContext, List dataList, List<String> extendHead) {
        this.excelContext = excelContext;
        dataMap = new HashMap<>(dataList == null ? 0 : dataList.size());
        if (CollUtil.isNotEmpty(dataList)) {
            for (int i = 0; i < dataList.size(); i++) {
                dataMap.put(i, dataList.get(i));
            }
        }
        this.extendHead = ListUtil.toList(extendHead);
    }

    @Override
    public void afterRowDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, Row row, Integer relativeRowIndex, Boolean isHead) {
        if (CollUtil.isEmpty(extendHead) || isNotClassHead || extendColumnField == null) {
            return;
        }
        if (Boolean.TRUE.equals(isHead)) {
            createExtendHead(writeSheetHolder, row);
            return;
        }
        writeExtendData(writeSheetHolder, row);
    }

    private void writeExtendData(WriteSheetHolder writeSheetHolder, Row row) {
        Object rawData = dataMap.get(row.getRowNum()-headRowNum);
        if (rawData == null) {
            return;
        }
        Map<String, String> extendData = getExtendData(rawData);
        if (CollUtil.isEmpty(extendData)) {
            return;
        }
        extendData.forEach((headName, data) -> {
            Integer columnIndex = extendHeadIndexMap.get(headName);
            if (columnIndex != null) {
                Cell cell = row.createCell(columnIndex);
                cell.setCellValue(data);
            }
        });
    }

    private void createExtendHead(WriteSheetHolder writeSheetHolder, Row row) {
        Map<Integer, Head> headMap = writeSheetHolder.excelWriteHeadProperty().getHeadMap();
        int rawColumnNum = headMap.size();
        int realColumnNum = rawColumnNum + extendHead.size();
        for (int i = rawColumnNum; i < realColumnNum; i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue(extendHead.get(i - rawColumnNum));
        }
    }

    @Override
    public void beforeSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {
        if (CollUtil.isEmpty(dataMap)) {
            return;
        }
        ExcelWriteHeadProperty excelWriteHeadProperty = writeSheetHolder.getExcelWriteHeadProperty();
        isNotClassHead = HeadKindEnum.CLASS != excelWriteHeadProperty.getHeadKind();
        headRowNum = excelWriteHeadProperty.getHeadRowNumber();

        List<FieldInfo> fieldInfos = ClassUtils.getClassFieldInfo(writeSheetHolder.getClazz());
        fieldInfos.stream().filter(FieldInfo::isExtendColumn)
                .findFirst().ifPresent(fieldInfo -> {
            extendColumnField = fieldInfo.getField();
            extendColumnField.setAccessible(true);
            fieldAccessor = new UnsafeFieldAccessor(extendColumnField);
            addHeadCache(excelWriteHeadProperty);
        });
        excelContext.setExtendHead(extendHead);
        excelContext.setHeadClass(writeSheetHolder.getClazz());
    }

    @Override
    public void afterSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {}

    private void addHeadCache(ExcelWriteHeadProperty excelWriteHeadProperty) {
        if (CollUtil.isNotEmpty(dataMap)) {
            Object rawData = dataMap.get(0);
            if (rawData != null) {
                Map<String, String> extendData = getExtendData(rawData);
                if (CollUtil.isNotEmpty(extendData)) {
                    extendHead.addAll(extendData.keySet());
                    extendHead = extendHead.stream().distinct()
                            .collect(Collectors.toList());
                }
            }
        }
        if (extendHead.size() == 0) {
            return;
        }
        int index = excelWriteHeadProperty.getHeadMap().size();
        for (String headName : extendHead) {
            extendHeadIndexMap.put(headName, index++);
        }
    }

    private Map<String, String> getExtendData(Object rawData) {
        return (Map<String, String>) fieldAccessor.getObject(rawData);
    }

}
