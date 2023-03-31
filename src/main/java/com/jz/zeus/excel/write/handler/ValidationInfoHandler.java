package com.jz.zeus.excel.write.handler;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.collection.CollectionUtil;
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
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;

import java.util.*;

/**
 * @Author JZ
 * @Date 2021/3/26 17:01
 */
public class ValidationInfoHandler extends AbstractSheetWriteHandler {

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
                validationInfoList.addAll(ClassUtils.getValidationInfos(writeSheetHolder.getClazz()));
            } else {
                ClassUtils.getValidationInfos(writeSheetHolder.getClazz())
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
            } else if (rowIndex != null) {
                if (columnIndex == null) {
                    addValidationData(workbook, sheet, rowIndex, rowIndex, 0, boxInfo.getColumnNum(), boxInfo);
                } else {
                    addValidationData(workbook, sheet, rowIndex, rowIndex, columnIndex, columnIndex, boxInfo);
                }
            } else if (boxInfo.isAsDicSheet()) {
                createValidationDataSheet(workbook, boxInfo);
            }
        }
    }

    /**
     * 为单元格生成下拉框
     * @param workbook     Excel所属workbook
     * @param sheet        数据所在sheet
     * @param firstRow     下拉框开始行索引
     * @param lastRow      下拉框结束行索引
     * @param firstCol     下拉框开始列索引
     * @param lastCol      下拉框结束列索引
     * @param boxInfo      下拉框配置信息
     */
    private void addValidationData(Workbook workbook, Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol, ValidationInfo boxInfo) {
        DataValidationHelper helper = sheet.getDataValidationHelper();
        DataValidationConstraint constraint;
        if (boxInfo.getParent() == null) {
            constraint = helper.createFormulaListConstraint(createValidationDataSheet(workbook, boxInfo));
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
//            if (parentBoxInfo.isAsDicSheet()) {
//                createValidationDataSheet()
//            }
            workbook.createSheet(parentSheetName);
        }
        String childSheetName = boxInfo.getSheetName();
        Sheet childSheet = Optional.ofNullable(workbook.getSheet(childSheetName))
                .orElse(workbook.createSheet(childSheetName));
        int rowIndex = 0;
        StrBuilder strBuilder = StrUtil.strBuilder();
        for (Map.Entry<String, List<String>> entry : boxInfo.getParentChildMap().entrySet()) {
            List<String> options = entry.getValue();
            if (CollectionUtil.isEmpty(options)) {
                continue;
            }
            for (int i = 0; i < options.size(); i++) {
                childSheet.createRow(rowIndex++).createCell(0)
                        .setCellValue(options.get(i));
            }
            // 拼接 name 格式：选项.父级sheet名.子集sheet名
            strBuilder.append(entry.getKey()).append(StrUtil.C_DOT)
                  .append(parentSheetName).append(StrUtil.C_DOT)
                  .append(childSheetName);
            // 创建一个 name
            ExcelUtils.createName(workbook, childSheetName, strBuilder.toStringAndReset(),
                  rowIndex - options.size() + 1, rowIndex, 0, 0);
        }

        String columnStr = ExcelUtils.columnIndexToStr(getColumnIndex(parentBoxInfo));
        for (int i = writeSheetHelper.getHeadRowNum(); i < boxInfo.getRowNum(); i++) {
            strBuilder.append(StrUtil.C_DOT).append(parentSheetName)
                  .append(StrUtil.C_DOT).append(childSheetName);
            String concatenate = String.format("ENCODEURL(CONCATENATE($%s$%d,\"%s\"))", columnStr, i+1, strBuilder.toStringAndReset());
            String substitute = "SUBSTITUTE(" + concatenate + ",\"%\",\"_\")";
            substitute = String.format("SUBSTITUTE(%s,\"-\",\"_2D\")", substitute);   // excel 中的 url encode 不对 '-' 进行处理，但 name 的名称不能包含 '-'，这里进行转换
            String formula = String.format("INDIRECT(%s)", substitute);
            DataValidationConstraint constraint = helper.createFormulaListConstraint(formula);
            CellRangeAddressList rangeAddressList = new CellRangeAddressList(i, i, firstCol, lastCol);
            DataValidation validation = createDataValidation(helper, constraint, rangeAddressList, boxInfo);
            sheet.addValidationData(validation);
        }

    }

    /**
     * 创建一个sheet，将下拉框的选项内容保存到该sheet中，并对数据区域创建一个名称
     * @param workbook
     * @param boxInfo
     * @return         数据区域的名称
     */
    private String createValidationDataSheet(Workbook workbook, ValidationInfo boxInfo) {
        int columnIndex = 0;
        int beginRowIndex = 0;
        String sheetName = boxInfo.getSheetName();
        Sheet sheet = Optional.ofNullable(workbook.getSheet(sheetName))
                .orElse(workbook.createSheet(sheetName));
        if (!boxInfo.isAsDicSheet()) {
            // 不作为字典表时将其设置到随机的位置
            columnIndex = RandomUtil.randomInt(200);
            beginRowIndex = RandomUtil.randomInt(1000);
            sheet.setColumnHidden(columnIndex, true);    // 将列进行隐藏
            sheet.protectSheet(IdUtil.fastSimpleUUID());   // 设置随机密码进行保护
            workbook.setSheetHidden(workbook.getSheetIndex(sheet), true);
        } else if (CharSequenceUtil.isNotBlank(boxInfo.getDicTitle())) {
            // 只有在作为字典表时，才能设置表头
            beginRowIndex = 1;
            Cell dicTitleCell = sheet.createRow(0).createCell(columnIndex);
            if (dicTitleCell instanceof HSSFCell) {
                dicTitleCell.setCellValue(new HSSFRichTextString(boxInfo.getDicTitle()));
            } else {
                dicTitleCell.setCellValue(boxInfo.getDicTitle());
            }
            CellStyle cellStyle = workbook.createCellStyle();
            Font font = workbook.createFont();
            font.setBold(true);
            font.setFontName("微软雅黑");
            cellStyle.setFont(font);
            dicTitleCell.setCellStyle(cellStyle);
            cellStyle.setWrapText(true);
            sheet.setColumnWidth(columnIndex, ExcelUtils.calColumnWidth(boxInfo.getDicTitle(), font.getFontHeightInPoints()));
        }
        List<String> options = boxInfo.getOptions();
        int endRowIndex = beginRowIndex + options.size() - 1;
        for (int i = beginRowIndex; i <= endRowIndex; i++) {
            sheet.createRow(i).createCell(columnIndex)
                    .setCellValue(options.get(i - beginRowIndex));
        }
        return ExcelUtils.createName(workbook, sheetName, StrUtil.C_UNDERLINE + RandomUtil.randomString(8),
              beginRowIndex + 1, endRowIndex + 1, columnIndex, columnIndex)
              .getNameName();
    }

    private DataValidation createDataValidation(DataValidationHelper helper, DataValidationConstraint constraint, CellRangeAddressList cellRangeAddressList, ValidationInfo boxInfo) {
        DataValidation dataValidation = helper.createValidation(constraint, cellRangeAddressList);
        if (dataValidation instanceof XSSFDataValidation) {
            dataValidation.setSuppressDropDownArrow(true);
        } else {
            dataValidation.setSuppressDropDownArrow(false);
        }
        if (boxInfo.isCheckDatavalidity()) {
            dataValidation.setShowErrorBox(true);
            dataValidation.createErrorBox(boxInfo.getErrorTitle(), boxInfo.getErrorMsg());
        }
        return dataValidation;
    }

    /**
     * 根据 ValidationInfo 获取列索引
     */
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

}
