package com.jz.zeus.excel.write.handler;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.util.IdUtil;
import cn.hutool.core.util.RandomUtil;
import cn.hutool.core.util.StrUtil;
import com.alibaba.excel.enums.HeadKindEnum;
import com.alibaba.excel.write.handler.AbstractSheetWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder;
import com.jz.zeus.excel.ValidationInfo;
import com.jz.zeus.excel.context.ExcelContext;
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

    /**
     * 下拉框超过该值后将写入到单独sheet中
     */
    private static final int MAX_GENERAL_OPTION_NUM = 20;

    private ExcelContext excelContext;

    private Integer headRowNum;

    private WriteSheetHelper writeSheetHelper;

    /**
     * 下拉框信息
     */
    private List<ValidationInfo> validationInfoList;

    public ValidationInfoHandler(ExcelContext excelContext, List<ValidationInfo> validationInfoList) {
        this(excelContext, null, validationInfoList);
    }

    public ValidationInfoHandler(ExcelContext excelContext, Integer headRowNum, List<ValidationInfo> validationInfoList) {
        this.excelContext = excelContext;
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
        writeSheetHelper = new WriteSheetHelper(excelContext, writeSheetHolder, headRowNum);

        Integer headRowNum = writeSheetHelper.getHeadRowNum();
        Workbook workbook = writeWorkbookHolder.getWorkbook();
        Sheet sheet = writeSheetHolder.getSheet();
        for (ValidationInfo boxInfo : validationInfoList) {
            if (CollUtil.isEmpty(boxInfo.getOptions())) {
                continue;
            }
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
                addValidationData(workbook, sheet, headRowNum, headRowNum+boxInfo.getRowNum(), columnIndex, columnIndex, boxInfo);
            } else if (rowIndex != null && columnIndex == null) {
                addValidationData(workbook, sheet, rowIndex, rowIndex, 0, boxInfo.getColumnNum(), boxInfo);
            } else if (rowIndex != null && columnIndex != null) {
                addValidationData(workbook, sheet, rowIndex, rowIndex, columnIndex, columnIndex, boxInfo);
            } else if (boxInfo.isAsDicSheet()) {
                addDictionarySheet(workbook, boxInfo);
            }
        }
    }

    private void addValidationData(Workbook workbook, Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol, ValidationInfo boxInfo) {
        List<String> options = boxInfo.getOptions();
        DataValidationHelper helper = sheet.getDataValidationHelper();
        DataValidationConstraint constraint;
        if (!boxInfo.isAsDicSheet() && options.size() <= MAX_GENERAL_OPTION_NUM) {
            constraint = helper.createExplicitListConstraint(options.toArray(new String[0]));
        } else {
            constraint = helper.createFormulaListConstraint(addHiddenValidationData(workbook, boxInfo));
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

    private String addHiddenValidationData(Workbook workbook, ValidationInfo boxInfo) {
        int columnIndex = 0;
        int beginRowIndex = 0;
        String sheetName = boxInfo.getSheetName();
        if (StrUtil.isBlank(sheetName)) {
            sheetName = "dic_".concat(RandomUtil.randomString(16));
        }
        Sheet sheet = workbook.createSheet(sheetName);
        sheet.protectSheet(IdUtil.fastSimpleUUID());
        if (!boxInfo.isAsDicSheet()) {
            columnIndex = RandomUtil.randomInt(200);
            beginRowIndex = RandomUtil.randomInt(1000);
            sheet.setColumnHidden(columnIndex, true);
            workbook.setSheetHidden(workbook.getSheetIndex(sheet), true);
        }
        List<String> options = boxInfo.getOptions();
        int endIndex = beginRowIndex + options.size() - 1;
        String s1 = null, s2 = null;
        for (int i = beginRowIndex; i <= endIndex; i++) {
            Cell newCell = sheet.createRow(i).createCell(columnIndex);
            if (i == beginRowIndex) {
                s1 = newCell.getAddress().formatAsString();
            }
            if (i == endIndex) {
                s2 = newCell.getAddress().formatAsString();
            }
            newCell.setCellValue(options.get(i - beginRowIndex));
        }
        Name categoryName = workbook.createName();
        categoryName.setNameName(sheetName);
        String refersToFormula = StrUtil.strBuilder().append(sheetName)
                .append("!$").append(StrUtil.filter(s1, Character::isLetter))
                .append('$').append(StrUtil.filter(s1, Character::isDigit))
                .append(":$").append(StrUtil.filter(s2, Character::isLetter))
                .append('$').append(StrUtil.filter(s2, Character::isDigit)).toString();
        categoryName.setRefersToFormula(refersToFormula);
        return sheetName;
    }

    private void addDictionarySheet(Workbook workbook, ValidationInfo boxInfo) {
        if (!boxInfo.isAsDicSheet()) {
            return;
        }
        String sheetName = boxInfo.getSheetName();
        if (StrUtil.isBlank(sheetName)) {
            sheetName = "dic_".concat(RandomUtil.randomString(16));
        }
        Sheet sheet = workbook.createSheet(sheetName);
        sheet.protectSheet(IdUtil.fastSimpleUUID());
        List<String> options = boxInfo.getOptions();
        for (int i = 0; i < options.size(); i++) {
            sheet.createRow(i).createCell(0).setCellValue(options.get(i));
        }
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
