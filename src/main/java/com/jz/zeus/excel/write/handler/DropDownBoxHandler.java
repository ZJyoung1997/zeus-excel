package com.jz.zeus.excel.write.handler;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.util.ArrayUtil;
import cn.hutool.core.util.StrUtil;
import com.alibaba.excel.metadata.CellData;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.write.handler.AbstractCellWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import com.jz.zeus.excel.DropDownBoxInfo;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;

import java.util.*;

/**
 * @Author JZ
 * @Date 2021/3/23 17:55
 */
@NoArgsConstructor
public class DropDownBoxHandler extends AbstractCellWriteHandler {

    private Map<String, Map<String, Integer>> excelHeadMap = new HashMap<>();

    private List<DropDownBoxInfo> dropDownBoxInfos = new ArrayList<>();

    public DropDownBoxHandler(List<DropDownBoxInfo> dropDownBoxInfos) {
        dropDownBoxInfos.forEach(info -> addDropDownBoxInfo(info));
    }

    @Override
    public void afterCellDataConverted(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, CellData cellData, Cell cell, Head head, Integer relativeRowIndex, Boolean isHead) {
        if (Boolean.TRUE.equals(isHead)) {
            Integer rowIndex = writeSheetHolder.getSheetNo();
            String sheetName = writeSheetHolder.getSheetName();
            String key = rowIndex == null ? "-" + sheetName : rowIndex + "-";
            Map<String, Integer> sheetHeadMap = excelHeadMap.get(key);
            if (CollUtil.isEmpty(sheetHeadMap)) {
                sheetHeadMap = new HashMap<>();
                excelHeadMap.put(key, sheetHeadMap);
            }
            sheetHeadMap.put(cell.getStringCellValue(), cell.getColumnIndex());
        }
    }

    @Override
    public void afterCellDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, List<CellData> cellDataList, Cell cell, Head head, Integer relativeRowIndex, Boolean isHead) {
        Sheet sheet = writeSheetHolder.getSheet();
        if (Boolean.TRUE.equals(isHead)) {
            Iterator<DropDownBoxInfo> iterator = dropDownBoxInfos.iterator();
            while (iterator.hasNext()) {
                DropDownBoxInfo info = iterator.next();
                if (info.getRowIndex() == null && (Integer.valueOf(cell.getColumnIndex()).equals(info.getColumnIndex()))
                        || cell.getStringCellValue().equals(info.getHeadName())) {
                    addValidationData(sheet, cell.getRowIndex()+1, cell.getRowIndex()+info.getRowNum(),
                            cell.getColumnIndex(), cell.getColumnIndex(), info.getOptions().toArray(new String[0]));
                    iterator.remove();
                } else if (info.getRowIndex() != null && info.getColumnIndex() == null
                            && StrUtil.isBlank(info.getHeadName())) {
                    addValidationData(sheet, info.getRowIndex(), info.getRowIndex(), cell.getColumnIndex()+1,
                            cell.getColumnIndex()+info.getColumnNum(), info.getOptions().toArray(new String[0]));
                    iterator.remove();
                }
            }
            return;
        }
        Iterator<DropDownBoxInfo> iterator = dropDownBoxInfos.iterator();
        while (iterator.hasNext()) {
            DropDownBoxInfo info = iterator.next();
            if (info.getRowIndex() != null) {
                if (info.getColumnIndex() != null) {
                    addValidationData(sheet, info.getRowIndex(), info.getRowIndex(), info.getColumnIndex(),
                            info.getColumnIndex(), info.getOptions().toArray(new String[0]));
                } else if (StrUtil.isNotBlank(info.getHeadName())) {
                    Integer headIndex = getHeadIndex(writeSheetHolder, info.getHeadName());
                    if (headIndex != null) {
                        addValidationData(sheet, info.getRowIndex(), info.getRowIndex(), headIndex,
                                headIndex, info.getOptions().toArray(new String[0]));
                    }
                }
                iterator.remove();
            }
        }

    }

    private void addValidationData(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol, String... options) {
        if (ArrayUtil.isEmpty(options)) {
            return;
        }
        DataValidationHelper helper = sheet.getDataValidationHelper();
        DataValidationConstraint constraint = helper.createExplicitListConstraint(options);
        DataValidation dataValidation = helper.createValidation(constraint,
                new CellRangeAddressList(firstRow, lastRow, firstCol, lastCol));
        if (dataValidation instanceof XSSFDataValidation) {
            dataValidation.setSuppressDropDownArrow(true);
            dataValidation.setShowErrorBox(true);
        } else {
            dataValidation.setSuppressDropDownArrow(false);
        }
        sheet.addValidationData(dataValidation);
    }

    private Integer getHeadIndex(WriteSheetHolder writeSheetHolder, String headName) {
        Integer rowIndex = writeSheetHolder.getSheetNo();
        String sheetName = writeSheetHolder.getSheetName();
        String key = rowIndex == null ? "-" + sheetName : rowIndex + "-";
        Map<String, Integer> sheetHeadMap = excelHeadMap.get(key);
        if (CollUtil.isEmpty(sheetHeadMap)) {
            return null;
        }
        return sheetHeadMap.get(headName);
    }

    public void addDropDownBoxInfo(DropDownBoxInfo boxInfo) {
        dropDownBoxInfos.add(boxInfo);
    }

}
