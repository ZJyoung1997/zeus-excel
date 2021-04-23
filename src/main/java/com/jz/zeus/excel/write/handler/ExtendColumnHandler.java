package com.jz.zeus.excel.write.handler;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.lang.Assert;
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
import java.util.*;

/**
 * 当存在动态表头时，该 handler应该最先被注册，已保证将动态表头添加到配置中
 * @Author JZ
 * @Date 2021/4/12 15:23
 */
public class ExtendColumnHandler extends AbstractRowWriteHandler implements SheetWriteHandler {

    private ExcelContext excelContext;

    private boolean isNotClassHead;

    private Map<String, Integer> extendHeadIndexMap = new HashMap<>();

    /**
     * 表头行数
     */
    private int headRowNum;

    /**
     * 待写入Excel数据，key 行索引、value 数据
     */
    private Map<Integer, Object> dataMap;

    /**
     * 扩展表头
     */
    private List<String> extendHead;

    /**
     * 扩展字段访问器
     */
    private UnsafeFieldAccessor fieldAccessor;

    public ExtendColumnHandler(ExcelContext excelContext) {
        Assert.notNull(excelContext, "ExcelContext must not be null");
        this.excelContext = excelContext;
    }

    @Override
    public void afterRowDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, Row row, Integer relativeRowIndex, Boolean isHead) {
        if (isNotClassHead || fieldAccessor == null || CollUtil.isEmpty(extendHead)) {
            return;
        }
        if (Boolean.TRUE.equals(isHead)) {
            writeExtendHead(writeSheetHolder, row);
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

    private void writeExtendHead(WriteSheetHolder writeSheetHolder, Row row) {
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
        ExcelWriteHeadProperty excelWriteHeadProperty = writeSheetHolder.getExcelWriteHeadProperty();
        isNotClassHead = HeadKindEnum.CLASS != excelWriteHeadProperty.getHeadKind();
        if (isNotClassHead) {
            return;
        }

        List dataList = excelContext.getSheetData();
        dataMap = new HashMap<>(dataList == null ? 0 : dataList.size());
        if (CollUtil.isNotEmpty(dataList)) {
            for (int i = 0; i < dataList.size(); i++) {
                dataMap.put(i, dataList.get(i));
            }
        }
        headRowNum = excelWriteHeadProperty.getHeadRowNumber();

        List<FieldInfo> fieldInfos = ClassUtils.getClassFieldInfo(writeSheetHolder.getClazz());
        fieldInfos.stream().filter(FieldInfo::isExtendColumn)
                .findFirst().ifPresent(fieldInfo -> {
            Field extendColumnField = fieldInfo.getField();
            extendColumnField.setAccessible(true);
            fieldAccessor = new UnsafeFieldAccessor(extendColumnField);
            addHeadCache(excelWriteHeadProperty);
        });
        excelContext.setExtendHead(extendHead);
    }

    @Override
    public void afterSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {}

    private void addHeadCache(ExcelWriteHeadProperty excelWriteHeadProperty) {
        if (CollUtil.isEmpty(dataMap) || fieldAccessor == null) {
            return;
        }
        Object rawData = null;
        for (Object value : dataMap.values()) {
            if (value != null) {
                rawData = value;
                break;
            }
        }
        if (rawData == null) {
            return;
        }
        Map<String, String> extendData = getExtendData(rawData);
        if (CollUtil.isNotEmpty(extendData)) {
            extendHead = new ArrayList<>(extendData.keySet());
        }
        if (CollUtil.isEmpty(extendHead)) {
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
