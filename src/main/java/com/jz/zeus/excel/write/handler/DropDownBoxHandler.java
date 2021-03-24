package com.jz.zeus.excel.write.handler;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.util.ArrayUtil;
import cn.hutool.core.util.StrUtil;
import com.alibaba.excel.metadata.CellData;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import com.jz.zeus.excel.DropDownBoxInfo;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * @Author JZ
 * @Date 2021/3/23 17:55
 */
@NoArgsConstructor
public class DropDownBoxHandler extends AbstractZeusCellWriteHandler {

    private List<DropDownBoxInfo> dropDownBoxInfos = new ArrayList<>();


    public DropDownBoxHandler(List<DropDownBoxInfo> dropDownBoxInfos) {
        if (CollUtil.isNotEmpty(dropDownBoxInfos)) {
            this.dropDownBoxInfos.addAll(dropDownBoxInfos);
        }
    }

    @Override
    public void afterCellDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, List<CellData> cellDataList, Cell cell, Head head, Integer relativeRowIndex, Boolean isHead) {
        if (CollUtil.isEmpty(dropDownBoxInfos)) {
            return;
        }
        if (!Boolean.TRUE.equals(isHead)) {
            return;
        }
        Sheet sheet = writeSheetHolder.getSheet();
        loadExcelHead(writeSheetHolder, cell);
        Iterator<DropDownBoxInfo> iterator = dropDownBoxInfos.iterator();
        while (iterator.hasNext()) {
            DropDownBoxInfo info = iterator.next();
            Integer boxColumnIndex = getHeadIndex(writeSheetHolder, info.getHeadName());
            if (info.getRowIndex() == null && (Integer.valueOf(cell.getColumnIndex()).equals(info.getColumnIndex())
                    || cell.getStringCellValue().equals(info.getHeadName()))) {
                addValidationData(sheet, cell.getRowIndex() + 1, cell.getRowIndex() + info.getRowNum(),
                        cell.getColumnIndex(), cell.getColumnIndex(), info.getOptions().toArray(new String[0]));
                iterator.remove();
            } else if (info.getRowIndex() != null) {
                if (info.getColumnIndex() == null && StrUtil.isBlank(info.getHeadName())) {
                    addValidationData(sheet, info.getRowIndex(), info.getRowIndex(), 0,
                            info.getColumnNum(), info.getOptions().toArray(new String[0]));
                    iterator.remove();
                } else if (info.getColumnIndex() != null || boxColumnIndex != null) {
                    Integer index = boxColumnIndex == null ? info.getColumnIndex() : boxColumnIndex;
                    addValidationData(sheet, info.getRowIndex(), info.getRowIndex(), index,
                            index, info.getOptions().toArray(new String[0]));
                    iterator.remove();
                }
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

    public DropDownBoxHandler addDropDownBoxInfo(DropDownBoxInfo boxInfo) {
        dropDownBoxInfos.add(boxInfo);
        return this;
    }

}
