package com.jz.zeus.excel.write.handler;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.util.ArrayUtil;
import cn.hutool.core.util.StrUtil;
import com.alibaba.excel.enums.HeadKindEnum;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder;
import com.jz.zeus.excel.ValidationInfo;
import com.jz.zeus.excel.util.ClassUtils;
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

/**
 * @Author JZ
 * @Date 2021/3/26 17:01
 */
public class ValidationInfoSheetHandler extends AbstractZeusSheetWriteHandler {

    /**
     * 下拉框信息
     */
    private List<ValidationInfo> validationInfoList;

    public ValidationInfoSheetHandler(List<ValidationInfo> validationInfoList) {
        this(null, validationInfoList);
    }

    public ValidationInfoSheetHandler(Integer headRowNum, List<ValidationInfo> validationInfoList) {
        super(headRowNum);
        this.validationInfoList = validationInfoList;
    }

    @Override
    public void afterSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {
        init(writeSheetHolder);
        List<ValidationInfo> boxInfoList = new ArrayList<>();
        if (CollUtil.isNotEmpty(validationInfoList)) {
            validationInfoList.forEach(boxInfo -> {
                if (boxInfo.getColumnIndex() == null && StrUtil.isNotBlank(boxInfo.getHeadName())) {
                    boxInfo.setColumnIndex(getHeadColumnIndex(boxInfo.getHeadName()));
                }
                boxInfoList.add(boxInfo);
            });
        }
        boolean isClass = HeadKindEnum.CLASS.equals(writeSheetHolder.getExcelWriteHeadProperty().getHeadKind());
        if (isClass) {
            if (CollUtil.isEmpty(validationInfoList)) {
                boxInfoList.addAll(ClassUtils.getValidationInfoInfos(writeSheetHolder.getClazz()));
            } else {
                ClassUtils.getValidationInfoInfos(writeSheetHolder.getClazz())
                        .forEach(boxInfo -> {
                            if (boxInfo.getColumnIndex() == null && StrUtil.isNotBlank(boxInfo.getHeadName())) {
                                boxInfo.setColumnIndex(getHeadColumnIndex(boxInfo.getHeadName()));
                            }
                            if (!validationInfoList.stream().filter(info -> Objects.equals(boxInfo, info)).findFirst().isPresent()) {
                                boxInfoList.add(boxInfo);
                            }
                        });
            }
        }

        if (CollUtil.isEmpty(boxInfoList)) {
            return;
        }

        Integer headRowNum = getHeadRowNum();
        Sheet sheet = writeSheetHolder.getSheet();
        boxInfoList.forEach(boxInfo -> {
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

    public void addValidationInfo(ValidationInfo validationInfo) {
        if (validationInfo != null) {
            if (CollUtil.isEmpty(validationInfoList)) {
                validationInfoList = new ArrayList<>();
            }
            validationInfoList.add(validationInfo);
        }
    }

    public void addValidationInfo(Integer[] columnIndexs, String... options) {
        if (ArrayUtil.isNotEmpty(columnIndexs) && ArrayUtil.isNotEmpty(options)) {
            if (CollUtil.isEmpty(validationInfoList)) {
                validationInfoList = new ArrayList<>();
            }
            List<String> optionList = Arrays.asList(options);
            for (int i = 0; i < columnIndexs.length; i++) {
                validationInfoList.add(ValidationInfo.buildColumn(columnIndexs[i], optionList));
            }
        }
    }

}
