package com.jz.zeus.excel.write.handler;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.util.StrUtil;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.write.handler.AbstractSheetWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.jz.zeus.excel.util.ExcelUtils;
import lombok.Getter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.HashMap;
import java.util.Map;

/**
 * @Author JZ
 * @Date 2021/3/26 17:11
 */
public class AbstractZeusSheetWriteHandler extends AbstractSheetWriteHandler {

    private static final Integer DEFAULT_HEAD_ROW_NUM = 1;

    /**
     * 表头行数，一般不用设置，调用initHeadRowNum()方法会自动判断
     */
    @Getter
    private Integer headRowNum;

    /**
     * key 表头、value 字段对应列索引
     */
    private Map<String, Integer> fieldIndexMap;


    public AbstractZeusSheetWriteHandler(Integer headRowNum) {
        this.headRowNum = headRowNum;
    }

    protected void init(WriteSheetHolder writeSheetHolder) {
        initHeadRowNum(writeSheetHolder);
        initFieldIndexMap(writeSheetHolder);
    }

    private void initHeadRowNum(WriteSheetHolder writeSheetHolder) {
        if (headRowNum == null) {
            headRowNum = writeSheetHolder.getExcelWriteHeadProperty().getHeadRowNumber();
            headRowNum = Math.max(headRowNum, DEFAULT_HEAD_ROW_NUM);
        }
    }

    private void initFieldIndexMap(WriteSheetHolder writeSheetHolder) {
        int lastRowNum = writeSheetHolder.getCachedSheet().getLastRowNum();
        Map<Integer, Head> headMap = writeSheetHolder.getExcelWriteHeadProperty().getHeadMap();
        if (CollUtil.isNotEmpty(headMap)) {
            headNameIndexMap = ExcelUtils.getHeadIndexMap(headMap);
        } else if (lastRowNum > 0) {
            if (headRowNum == null) {
                initHeadRowNum(writeSheetHolder);
            }
            Sheet sheet = writeSheetHolder.getCachedSheet();
            headNameIndexMap = new HashMap<>();
            for (int i = 0; i < headRowNum; i++) {
                Row headRow = sheet.getRow(i);
                int cellNum = headRow.getLastCellNum();
                for (int j = 0; j < cellNum; j++) {
                    Cell cell = headRow.getCell(j);
                    if (cell == null || StrUtil.isBlank(cell.getStringCellValue())) {
                        continue;
                    }
                    headNameIndexMap.put(cell.getStringCellValue(), j);
                }
            }
        } else if (headRowNum == null) {
            initHeadRowNum(writeSheetHolder);
        }
    }

    protected Integer getHeadColumnIndex(String headName) {
        return CollUtil.isEmpty(headNameIndexMap) ? null : headNameIndexMap.get(headName);
    }

}
