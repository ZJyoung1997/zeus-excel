package com.jz.zeus.excel.write.handler;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.util.ArrayUtil;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder;
import com.jz.zeus.excel.DropDownBoxInfo;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;

import java.util.Arrays;
import java.util.List;

/**
 * @Author JZ
 * @Date 2021/3/26 17:01
 */
public class DropDownBoxSheetHandler extends AbstractZeusSheetWriteHandler {

    /**
     * 下拉框信息
     */
    private List<DropDownBoxInfo> dropDownBoxInfoList;

    public DropDownBoxSheetHandler(List<DropDownBoxInfo> dropDownBoxInfoList) {
        this(null, dropDownBoxInfoList);
    }

    public DropDownBoxSheetHandler(Integer headRowNum, List<DropDownBoxInfo> dropDownBoxInfoList) {
        super(headRowNum);
        this.dropDownBoxInfoList = dropDownBoxInfoList;
    }

    @Override
    public void afterSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {
        if (CollUtil.isEmpty(dropDownBoxInfoList)) {
            return;
        }

        initHeadRowNum(writeSheetHolder);
        initHeadNameIndexMap(writeSheetHolder);

        Integer headRowNum = getHeadRowNum();
        Sheet sheet = writeSheetHolder.getSheet();
        dropDownBoxInfoList.forEach(boxInfo -> {
            Integer rowIndex = boxInfo.getRowIndex();
            Integer columnIndex = boxInfo.getColumnIndex();
            if (columnIndex == null) {
                columnIndex = getHeadColumnIndex(boxInfo.getHeadName());
            }
            if (rowIndex == null && columnIndex != null) {
                addValidationData(sheet, headRowNum, headRowNum+boxInfo.getRowNum(), columnIndex, columnIndex, boxInfo.getOptions());
            } else if (rowIndex != null && columnIndex == null) {
                addValidationData(sheet, rowIndex, rowIndex, 0, boxInfo.getColumnNum(), boxInfo.getOptions());
            } else if (rowIndex != null && columnIndex != null) {
                addValidationData(sheet, rowIndex, rowIndex, columnIndex, columnIndex, boxInfo.getOptions());
            }
        });
    }

    private void addValidationData(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol, List<String> options) {
        if (CollUtil.isEmpty(options)) {
            return;
        }
        DataValidationHelper helper = sheet.getDataValidationHelper();
        DataValidationConstraint constraint = helper.createExplicitListConstraint(options.toArray(new String[0]));
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

    public void addDropDownBoxInfo(DropDownBoxInfo dropDownBoxInfo) {
        if (dropDownBoxInfo != null) {
            dropDownBoxInfoList.add(dropDownBoxInfo);
        }
    }

    public void addDropDownBoxInfo(Integer[] columnIndexs, String... options) {
        if (ArrayUtil.isNotEmpty(columnIndexs) && ArrayUtil.isNotEmpty(options)) {
            List<String> optionList = Arrays.asList(options);
            for (int i = 0; i < columnIndexs.length; i++) {
                dropDownBoxInfoList.add(new DropDownBoxInfo(columnIndexs[i], optionList));
            }
        }
    }

}
