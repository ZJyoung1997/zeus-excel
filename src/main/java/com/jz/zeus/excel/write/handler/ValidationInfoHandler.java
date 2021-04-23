package com.jz.zeus.excel.write.handler;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.lang.Filter;
import cn.hutool.core.util.StrUtil;
import com.alibaba.excel.enums.HeadKindEnum;
import com.alibaba.excel.write.handler.AbstractSheetWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder;
import com.jz.zeus.excel.ValidationInfo;
import com.jz.zeus.excel.util.ClassUtils;
import com.jz.zeus.excel.write.helper.WriteSheetHelper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/**
 * @Author JZ
 * @Date 2021/3/26 17:01
 */
public class ValidationInfoHandler extends AbstractSheetWriteHandler {

    private Integer headRowNum;

    private WriteSheetHelper writeSheetHelper;

    /**
     * 下拉框信息
     */
    private List<ValidationInfo> validationInfoList;

    public ValidationInfoHandler(List<ValidationInfo> validationInfoList) {
        this(null, validationInfoList);
    }

    public ValidationInfoHandler(Integer headRowNum, List<ValidationInfo> validationInfoList) {
        this.headRowNum = headRowNum;
        this.validationInfoList = validationInfoList;
    }

    @Override
    public void afterSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {
        boolean isClass = HeadKindEnum.CLASS == writeSheetHolder.getExcelWriteHeadProperty().getHeadKind();
        if (CollUtil.isEmpty(validationInfoList)) {
            validationInfoList = new ArrayList<>();
        }
        if (isClass) {
            if (CollUtil.isEmpty(validationInfoList)) {
                validationInfoList.addAll(ClassUtils.getValidationInfoInfos(writeSheetHolder.getClazz()));
            } else {
                ClassUtils.getValidationInfoInfos(writeSheetHolder.getClazz())
                        .forEach(boxInfo -> {
                            if (validationInfoList.stream().noneMatch(info -> Objects.equals(boxInfo, info))) {
                                validationInfoList.add(boxInfo);
                            }
                        });
            }
        }
        if (CollUtil.isEmpty(validationInfoList)) {
            return;
        }
        writeSheetHelper = new WriteSheetHelper(writeSheetHolder, headRowNum);

        Integer headRowNum = writeSheetHelper.getHeadRowNum();
        Workbook workbook = writeWorkbookHolder.getWorkbook();
        Sheet sheet = writeSheetHolder.getSheet();
        int length = validationInfoList.size();
        for (int i = 0; i < length; i++) {
            ValidationInfo boxInfo = validationInfoList.get(i);
            Integer rowIndex = boxInfo.getRowIndex();
            Integer columnIndex = boxInfo.getColumnIndex();
            if (columnIndex == null) {
                if (StrUtil.isNotBlank(boxInfo.getFieldName())) {
                    columnIndex = writeSheetHelper.getFieldColumnIndex(boxInfo.getFieldName());
                } else if (StrUtil.isNotBlank(boxInfo.getHeadName())) {
                    columnIndex = writeSheetHelper.getHeadColumnIndex(boxInfo.getHeadName());
                }
            }
            if (rowIndex == null && columnIndex != null) {
                addValidationData(workbook, sheet, headRowNum, headRowNum+boxInfo.getRowNum(), columnIndex, columnIndex, boxInfo.getOptions(), i);
            } else if (rowIndex != null && columnIndex == null) {
                addValidationData(workbook, sheet, rowIndex, rowIndex, 0, boxInfo.getColumnNum(), boxInfo.getOptions(), i);
            } else if (rowIndex != null && columnIndex != null) {
                addValidationData(workbook, sheet, rowIndex, rowIndex, columnIndex, columnIndex, boxInfo.getOptions(), i);
            }
        }
    }

    private void addValidationData(Workbook workbook, Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol, List<String> options, int index) {
        if (CollUtil.isEmpty(options)) {
            return;
        }

        DataValidationHelper helper = sheet.getDataValidationHelper();
        DataValidationConstraint constraint = null;
        if (options.size() > 10) {
            constraint = helper.createFormulaListConstraint(addHiddenValidationData(workbook, firstCol, options, index));
        } else {
            constraint = helper.createExplicitListConstraint(options.toArray(new String[0]));
        }
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

    private String addHiddenValidationData(Workbook workbook, Integer columnIndex, List<String> options, int index) {
        String hiddenSheetName = "hidden".concat(String.valueOf(System.currentTimeMillis() + index));
        Sheet hiddenSheet = workbook.createSheet(hiddenSheetName);
        workbook.setSheetHidden(workbook.getSheetIndex(hiddenSheet), true);
        int optionsSize = options.size();
        int endIndex = optionsSize - 1;
        String s1 = null, s2 = null;
        for (int i = 0; i < optionsSize; i++) {
            Cell newCell = hiddenSheet.createRow(i).createCell(columnIndex);
            if (i == 0) {
                s1 = newCell.getAddress().formatAsString();
            }
            if (i == endIndex) {
                s2 = newCell.getAddress().formatAsString();
            }
            newCell.setCellValue(options.get(i));
        }
        Name categoryName = workbook.createName();
        categoryName.setNameName(hiddenSheetName);
        String refersToFormula = StrUtil.builder().append(hiddenSheetName)
                .append("!$").append(StrUtil.filter(s1, Character::isLetter))
                .append('$').append(StrUtil.filter(s1, Character::isDigit))
                .append(":$").append(StrUtil.filter(s2, Character::isLetter))
                .append('$').append(StrUtil.filter(s2, Character::isDigit)).toString();
        categoryName.setRefersToFormula(refersToFormula);
        return hiddenSheetName;
    }

    public void addValidationInfo(ValidationInfo validationInfo) {
        if (validationInfo == null) {
            return;
        }
        if (CollUtil.isEmpty(validationInfoList)) {
            validationInfoList = new ArrayList<>();
        }
        validationInfoList.add(validationInfo);
    }

    public void addValidationInfo(Integer[] columnIndexs, String... options) {
        validationInfoList.addAll(ValidationInfo.buildCommonOption(columnIndexs, options));
    }

}
