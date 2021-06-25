package com.jz.zeus.excel.write.helper;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.lang.Assert;
import cn.hutool.core.map.MapUtil;
import cn.hutool.core.util.StrUtil;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.property.ExcelWriteHeadProperty;
import com.jz.zeus.excel.context.ExcelContext;
import lombok.Getter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.HashMap;
import java.util.Map;

/**
 * @Author JZ
 * @Date 2021/4/16 16:09
 */
public class WriteSheetHelper {

    private static final Integer DEFAULT_HEAD_ROW_NUM = 1;

    private ExcelContext excelContext;

    /**
     * 表头行数，一般不用设置，调用initHeadRowNum()方法会自动判断
     */
    @Getter
    private Integer headRowNum;

    /**
     * key 属性名、value 属性对应列索引
     */
    private Map<String, Integer> fieldNameIndexMap;

    /**
     * key 表头、value 表头对应列索引
     */
    private Map<String, Integer> headNameIndexMap;

    private WriteSheetHolder writeSheetHolder;


    public WriteSheetHelper(ExcelContext excelContext, WriteSheetHolder writeSheetHolder) {
        this(excelContext, writeSheetHolder, null);
    }


    public WriteSheetHelper(ExcelContext excelContext, WriteSheetHolder writeSheetHolder, Integer headRowNum) {
        Assert.notNull(writeSheetHolder, "WriteSheetHolder can't be null");
        this.excelContext = excelContext;
        this.headRowNum = headRowNum;
        this.writeSheetHolder = writeSheetHolder;
        initCache();
    }

    private void initHeadRowNum() {
        if (headRowNum == null) {
            headRowNum = writeSheetHolder.getExcelWriteHeadProperty().getHeadRowNumber();
            headRowNum = Math.max(headRowNum, DEFAULT_HEAD_ROW_NUM);
        }
    }

    private void initCache() {
        int lastRowNum = writeSheetHolder.getCachedSheet().getLastRowNum();
        ExcelWriteHeadProperty excelWriteHeadProperty = writeSheetHolder.getExcelWriteHeadProperty();
        Map<Integer, Head> headMap = excelWriteHeadProperty.getHeadMap();
        if (headNameIndexMap == null) {
            headNameIndexMap = new HashMap<>(headMap.size());
        } else {
            headNameIndexMap.clear();
        }
        if (fieldNameIndexMap == null) {
            fieldNameIndexMap = new HashMap<>(headMap.size());
        } else {
            fieldNameIndexMap.clear();
        }
        initHeadRowNum();

        if (CollUtil.isNotEmpty(headMap)) {
            headMap.values().forEach(head -> {
                Integer columnIndex = head.getColumnIndex();
                head.getHeadNameList().forEach(headName -> {
                    headNameIndexMap.put(headName, columnIndex);
                });
                if (StrUtil.isNotBlank(head.getFieldName())) {
                    fieldNameIndexMap.put(head.getFieldName(), columnIndex);
                }
            });
        } else if (lastRowNum > 0) {
            Sheet sheet = writeSheetHolder.getCachedSheet();
            for (int i = 0; i < headRowNum; i++) {
                Row headRow = sheet.getRow(i);
                int cellNum = headRow.getLastCellNum();
                for (int j = 0; j < cellNum; j++) {
                    Cell cell = headRow.getCell(j);
                    String cellValue;
                    if (cell == null || StrUtil.isBlank((cellValue = cell.getStringCellValue()))) {
                        continue;
                    }
                    if (!headNameIndexMap.keySet().contains(cellValue)) {
                        headNameIndexMap.put(cellValue, j);
                    }
                }
            }
        }

        if (CollUtil.isNotEmpty(excelContext.getExtendHead())) {
            headNameIndexMap.putAll(excelContext.getExtendHead());
        }
        if (MapUtil.isNotEmpty(excelContext.getDynamicHead())) {
            headNameIndexMap.putAll(excelContext.getDynamicHead());
        }
    }

    public Integer getHeadColumnIndex(String headName) {
        return CollUtil.isEmpty(headNameIndexMap) ? null : headNameIndexMap.get(headName);
    }

    public Integer getFieldColumnIndex(String fieldName) {
        return CollUtil.isEmpty(fieldNameIndexMap) ? null : fieldNameIndexMap.get(fieldName);
    }

}
