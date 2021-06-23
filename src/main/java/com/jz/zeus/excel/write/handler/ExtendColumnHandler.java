package com.jz.zeus.excel.write.handler;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.lang.Assert;
import cn.hutool.core.util.ReflectUtil;
import com.alibaba.excel.enums.HeadKindEnum;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder;
import com.alibaba.excel.write.property.ExcelWriteHeadProperty;
import com.jz.zeus.excel.FieldInfo;
import com.jz.zeus.excel.context.ExcelContext;
import com.jz.zeus.excel.util.ClassUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 当存在动态表头时，该 handler应该最先被注册，已保证将动态表头添加到配置中
 *
 * @Author JZ
 * @Date 2021/4/12 15:23
 */
public class ExtendColumnHandler extends AbstractRowWriteHandler {

    private ExcelContext excelContext;

    private boolean isNotClassHead;

    private Map<String, Integer> extendHeadIndexMap = new HashMap<>();

    /**
     * 待写入Excel数据
     */
    private List dataList;

    /**
     * 扩展表头
     */
    private List<String> extendHead;

    /**
     * 扩展字段
     */
    private Field extendColumnField;

    /**
     * 写入数据的起始行索引
     */
    private Integer writeDataBeginRowIndex;

    public ExtendColumnHandler(ExcelContext excelContext) {
        Assert.notNull(excelContext, "ExcelContext must not be null");
        this.excelContext = excelContext;
    }

    @Override
    public void afterRowCreate(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, Row row, Integer relativeRowIndex, Boolean isHead) {
        if (isNotClassHead || CollUtil.isEmpty(extendHead)) {
            return;
        }
        if (Boolean.TRUE.equals(isHead)) {
            writeExtendHead(writeSheetHolder, row);
            return;
        }
        if (writeDataBeginRowIndex == null) {
            writeDataBeginRowIndex = row.getRowNum();
        }
        writeExtendData(writeSheetHolder, row);
    }

    private void writeExtendData(WriteSheetHolder writeSheetHolder, Row row) {
        Object rawData = dataList.get(row.getRowNum() - writeDataBeginRowIndex);
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

        dataList = excelContext.getSheetData();

        extendColumnField = null;
        List<FieldInfo> fieldInfos = ClassUtils.getClassFieldInfo(writeSheetHolder.getClazz());
        fieldInfos.stream().filter(FieldInfo::isExtendColumn)
                .findFirst().ifPresent(fieldInfo -> {
            extendColumnField = fieldInfo.getField();
            addHeadCache(excelWriteHeadProperty);
        });
        if (extendColumnField == null) {
            isNotClassHead = true;
            return;
        }
        excelContext.setExtendHead(new HashMap<>(extendHeadIndexMap));
    }

    private void addHeadCache(ExcelWriteHeadProperty excelWriteHeadProperty) {
        if (CollUtil.isEmpty(dataList) || extendColumnField == null) {
            return;
        }
        Object rawData = null;
        for (Object value : dataList) {
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
        return (Map<String, String>) ReflectUtil.getFieldValue(rawData, extendColumnField);
    }

}
