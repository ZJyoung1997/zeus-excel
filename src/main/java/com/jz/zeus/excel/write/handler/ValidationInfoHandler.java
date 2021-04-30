package com.jz.zeus.excel.write.handler;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.text.CharSequenceUtil;
import cn.hutool.core.text.StrBuilder;
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
import com.jz.zeus.excel.util.ExcelUtils;
import com.jz.zeus.excel.write.helper.WriteSheetHelper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;

import java.util.*;

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
            if (boxInfo.getParent() == null && CollUtil.isEmpty(boxInfo.getOptions())) {
                continue;
            }
            Integer rowIndex = boxInfo.getRowIndex();
            Integer columnIndex = getColumnIndex(boxInfo);
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
        if (boxInfo.getParent() == null) {
            if (!boxInfo.isAsDicSheet() && options.size() <= MAX_GENERAL_OPTION_NUM) {
                constraint = helper.createExplicitListConstraint(options.toArray(new String[0]));
            } else {
                constraint = helper.createFormulaListConstraint(addHiddenValidationData(workbook, boxInfo));
            }
            DataValidation dataValidation = createDataValidation(helper, constraint,
                    new CellRangeAddressList(firstRow, lastRow, firstCol, lastCol), boxInfo);
            sheet.addValidationData(dataValidation);
        } else {
            addCascadeValidationData(workbook, sheet, helper, boxInfo, firstCol, lastCol);
        }
    }

    /**
     * 添加级联下拉框
     */
    private void addCascadeValidationData(Workbook workbook, Sheet sheet, DataValidationHelper helper, ValidationInfo boxInfo, int firstCol, int lastCol) {
        ValidationInfo parentBoxInfo = boxInfo.getParent();
        String parentSheetName = parentBoxInfo.getSheetName();
        if (workbook.getSheet(parentSheetName) == null) {
            workbook.createSheet(parentSheetName);
        }
        String childCheetName = boxInfo.getSheetName();
        Sheet childSheet = Optional.ofNullable(workbook.getSheet(childCheetName))
                .orElse(workbook.createSheet(childCheetName));
        int rowIndex = 0;
        StrBuilder strBuilder = StrUtil.strBuilder();
        for (Map.Entry<String, List<String>> entry : boxInfo.getParentChildMap().entrySet()) {
            List<String> siteList = entry.getValue();
            String s1 = null, s2 = null;
            for (int i = 0; i < siteList.size(); i++) {
                Cell cell = childSheet.createRow(rowIndex++).createCell(0);
                cell.setCellValue(siteList.get(i));
                if (i == 0) {
                    s1 = cell.getAddress().formatAsString();
                }
                if (i == siteList.size() - 1) {
                    s2 = cell.getAddress().formatAsString();
                }
            }

            String parent = entry.getKey() + parentSheetName;
            if (workbook.getName(parent) != null) {
                continue;
            }
            Name categoryName = workbook.createName();
            categoryName.setNameName(parent);
            String refersToFormula = strBuilder.append(childCheetName)
                    .append("!$").append(StrUtil.filter(s1, Character::isLetter))
                    .append('$').append(StrUtil.filter(s1, Character::isDigit))
                    .append(":$").append(StrUtil.filter(s2, Character::isLetter))
                    .append('$').append(StrUtil.filter(s2, Character::isDigit)).toString();
            categoryName.setRefersToFormula(refersToFormula);
            strBuilder.reset();
        }

        StrBuilder formulaBuild = StrUtil.strBuilder();
        String columnStr = ExcelUtils.columnIndexToStr(getColumnIndex(parentBoxInfo));
        for (int i = writeSheetHelper.getHeadRowNum(); i < boxInfo.getRowNum(); i++) {
            formulaBuild.append("INDIRECT(CONCAT($").append(columnStr).append('$').append(i + 1)
                    .append(",\"").append(parentSheetName).append("\"))");
            DataValidationConstraint constraint = helper.createFormulaListConstraint(formulaBuild.toString());
            CellRangeAddressList rangeAddressList = new CellRangeAddressList(i, i, firstCol, lastCol);
            DataValidation validation = createDataValidation(helper, constraint, rangeAddressList, boxInfo);
            sheet.addValidationData(validation);
            formulaBuild.reset();
        }

    }

    private String addHiddenValidationData(Workbook workbook, ValidationInfo boxInfo) {
        int columnIndex = 0;
        int beginRowIndex = 0;
        String sheetName = boxInfo.getSheetName();
        Sheet sheet = Optional.ofNullable(workbook.getSheet(sheetName))
                .orElse(workbook.createSheet(sheetName));
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
        categoryName.setNameName(StrUtil.C_UNDERLINE + RandomUtil.randomString(8));
        String refersToFormula = StrUtil.strBuilder().append(sheetName)
                .append("!$").append(StrUtil.filter(s1, Character::isLetter))
                .append('$').append(StrUtil.filter(s1, Character::isDigit))
                .append(":$").append(StrUtil.filter(s2, Character::isLetter))
                .append('$').append(StrUtil.filter(s2, Character::isDigit)).toString();
        categoryName.setRefersToFormula(refersToFormula);
        return categoryName.getNameName();
    }

    /**
     * 将下拉框信息作为字典表，且不会创建下拉框
     */
    private void addDictionarySheet(Workbook workbook, ValidationInfo boxInfo) {
        String sheetName = boxInfo.getSheetName();
        Sheet sheet = workbook.createSheet(sheetName);
        sheet.protectSheet(IdUtil.fastSimpleUUID());
        List<String> options = boxInfo.getOptions();
        for (int i = 0; i < options.size(); i++) {
            sheet.createRow(i).createCell(0).setCellValue(options.get(i));
        }
    }

    private DataValidation createDataValidation(DataValidationHelper helper, DataValidationConstraint constraint, CellRangeAddressList cellRangeAddressList, ValidationInfo boxInfo) {
        DataValidation dataValidation = helper.createValidation(constraint, cellRangeAddressList);
        if (dataValidation instanceof XSSFDataValidation) {
            dataValidation.setShowErrorBox(true);
            dataValidation.createErrorBox(boxInfo.getErrorTitle(), boxInfo.getErrorMsg());
            dataValidation.setSuppressDropDownArrow(true);
        } else {
            dataValidation.setSuppressDropDownArrow(false);
        }
        return dataValidation;
    }

    private Integer getColumnIndex(ValidationInfo boxInfo) {
        Integer columnIndex = boxInfo.getColumnIndex();
        if (columnIndex == null) {
            if (StrUtil.isNotBlank(boxInfo.getFieldName())) {
                columnIndex = writeSheetHelper.getFieldColumnIndex(boxInfo.getFieldName());
            } else if (StrUtil.isNotBlank(boxInfo.getHeadName())) {
                columnIndex = writeSheetHelper.getHeadColumnIndex(boxInfo.getHeadName());
            }
        }
        return columnIndex;
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
