package com.jz.zeus.excel.write.handler;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.util.ArrayUtil;
import cn.hutool.core.util.StrUtil;
import com.alibaba.excel.enums.HeadKindEnum;
import com.alibaba.excel.write.handler.AbstractSheetWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder;
import com.jz.zeus.excel.ValidationInfo;
import com.jz.zeus.excel.util.ClassUtils;
import com.jz.zeus.excel.write.helper.WriteSheetHelper;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Objects;
import java.util.stream.Collectors;

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
                            List<Object> li = validationInfoList.stream().filter(info -> Objects.equals(boxInfo, info)).collect(Collectors.toList());
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
        Sheet sheet = writeSheetHolder.getSheet();
        validationInfoList.forEach(boxInfo -> {
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
