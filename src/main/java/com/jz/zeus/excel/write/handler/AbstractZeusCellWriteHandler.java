package com.jz.zeus.excel.write.handler;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.util.StrUtil;
import com.alibaba.excel.write.handler.AbstractCellWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import org.apache.poi.ss.usermodel.Cell;

import java.util.*;

/**
 * @Author JZ
 * @Date 2021/3/24 10:36
 */
public abstract class AbstractZeusCellWriteHandler extends AbstractCellWriteHandler {

    protected Map<String, List<String>> excelHeadMap;

    protected void loadExcelHead(WriteSheetHolder writeSheetHolder, Cell cell) {
        if (CollUtil.isEmpty(excelHeadMap)) {
            excelHeadMap = new HashMap<>();
        }
        Integer rowIndex = writeSheetHolder.getSheetNo();
        String sheetName = writeSheetHolder.getSheetName();
        String key = rowIndex == null ? "-" + sheetName : rowIndex + "-";
        List<String> sheetHeadMap = excelHeadMap.get(key);
        if (CollUtil.isEmpty(sheetHeadMap)) {
            sheetHeadMap = new ArrayList<>();
            excelHeadMap.put(key, sheetHeadMap);
        }
        sheetHeadMap.add(cell.getColumnIndex(), cell.getStringCellValue());
    }

    protected String getHeadName(WriteSheetHolder writeSheetHolder, Integer index) {
        if (index == null || index < 0) {
            return null;
        }
        List<String> sheetHead = getCurrentSheetHead(writeSheetHolder.getSheetNo(), writeSheetHolder.getSheetName());
        return index >= sheetHead.size() ? null : sheetHead.get(index);
    }

    protected Integer getHeadIndex(WriteSheetHolder writeSheetHolder, String headName) {
        if (StrUtil.isBlank(headName)) {
            return null;
        }
        List<String> sheetHeads = getCurrentSheetHead(writeSheetHolder.getSheetNo(), writeSheetHolder.getSheetName());
        for (int i = 0; i < sheetHeads.size(); i++) {
            if (headName.equals(sheetHeads.get(i))) {
                return i;
            }
        }
        return null;
    }

    protected List<String> getCurrentSheetHead(Integer rowIndex, String sheetName) {
        if (CollUtil.isEmpty(excelHeadMap)) {
            return Collections.emptyList();
        }
        String key = rowIndex == null ? "-" + sheetName : rowIndex + "-";
        List<String> result = excelHeadMap.get(key);
        return result == null ? Collections.emptyList() : result;
    }

}
