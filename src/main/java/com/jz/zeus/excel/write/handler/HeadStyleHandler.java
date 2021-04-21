package com.jz.zeus.excel.write.handler;

import cn.hutool.core.bean.BeanUtil;
import cn.hutool.core.collection.CollUtil;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.write.handler.AbstractRowWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import com.alibaba.excel.write.property.ExcelWriteHeadProperty;
import com.jz.zeus.excel.FieldInfo;
import com.jz.zeus.excel.constant.Constants;
import com.jz.zeus.excel.context.ExcelContext;
import com.jz.zeus.excel.util.ClassUtils;
import com.jz.zeus.excel.util.StringUtils;
import com.jz.zeus.excel.write.property.CellStyleProperty;
import org.apache.poi.ss.usermodel.*;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * @Author JZ
 * @Date 2021/3/23 11:57
 */
public class HeadStyleHandler extends AbstractRowWriteHandler {

    /**
     * 是否自适应列宽，默认开启
     */
    private boolean isAutoColumnWidth = true;

    /**
     * 是否开启无边框
     * 由于 EasyExcel 使用注解设置表头样式时默认有边界，
     * 这里为true，表头将无边框，仅在使用注解配置表头样式时生效
     */
    private boolean enabledNoBorder = true;

    /**
     * 表头单元格样式，外层list下标对应行索引，内层下标对应列索引
     */
    private List<List<CellStyleProperty>> multiRowHeadCellStyles;

    private CellStyleProperty allHeadStyle;

    /**
     * 若使用class表示表头且指定样式时，优先使用自定义样式，若无则使用class指定样式，若没有指定任何样式则使用默认样式。
     * 若是用list自定义表头则赋予默认表头样式
     */
    public HeadStyleHandler() {}

    /**
     * 为所有设置表头同一样式
     */
    public HeadStyleHandler(CellStyleProperty allHeadStyle) {
        this.allHeadStyle = allHeadStyle;
    }

    /**
     * 设置行索引为 0 的表头的样式
     */
    public HeadStyleHandler(List<CellStyleProperty> singleRowHeadCellStyles) {
        if (CollUtil.isNotEmpty(singleRowHeadCellStyles)) {
            this.multiRowHeadCellStyles = new ArrayList<>(1);
            this.multiRowHeadCellStyles.add(singleRowHeadCellStyles);
        }
    }

    /**
     * 对表头的每个单元格按指定样式进行设置
     * @param multiRowHeadCellStyles  表头单元格样式，外层list下标对应行索引，内层下标对应列索引
     */
    public HeadStyleHandler setMultiRowHeadCellStyles(List<List<CellStyleProperty>> multiRowHeadCellStyles) {
        this.multiRowHeadCellStyles = multiRowHeadCellStyles;
        return this;
    }

    @Override
    public void afterRowDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, Row row, Integer relativeRowIndex, Boolean isHead) {
        if (!Boolean.TRUE.equals(isHead)) {
            return;
        }
        ExcelWriteHeadProperty excelWriteHeadProperty = writeSheetHolder.getExcelWriteHeadProperty();
        Class headClass = excelWriteHeadProperty.getClass();
        Map<Integer, Head> headMap = excelWriteHeadProperty.getHeadMap();
        Sheet sheet = writeSheetHolder.getSheet();
        int rawColumnNum = headMap.size();
        int realColumnNum = rawColumnNum + (CollUtil.isEmpty(ExcelContext.getExtendHead()) || headClass != ExcelContext.getHeadClass() ?
                0 : ExcelContext.getExtendHead().size());
        for (int i = 0; i < realColumnNum; i++) {
            Cell cell = row.getCell(i);
            if (cell == null) {
                continue;
            }
            boolean isSet = false;
            if (allHeadStyle != null) {
                setCellStyle(sheet, cell, allHeadStyle);
                isSet = true;
            } else if (CollUtil.isNotEmpty(multiRowHeadCellStyles) && cell.getRowIndex() < multiRowHeadCellStyles.size()) {
                List<CellStyleProperty> cellStylePropertyList = multiRowHeadCellStyles.get(cell.getRowIndex());
                if (CollUtil.isNotEmpty(cellStylePropertyList) && cell.getColumnIndex() < cellStylePropertyList.size()) {
                    setCellStyle(sheet, cell, cellStylePropertyList.get(i));
                    isSet = true;
                }
            }
            if (!isSet) {
                if (i >= rawColumnNum) {
                    CellStyleProperty styleProperty = CellStyleProperty.getDefaultHeadProperty();
                    ClassUtils.getClassFieldInfo(headClass).stream()
                            .filter(FieldInfo::isExtendColumn)
                            .findFirst()
                            .ifPresent(fieldInfo -> {
                                if (fieldInfo.getHeadFontProperty() != null) {
                                    styleProperty.setFontProperty(fieldInfo.getHeadFontProperty());
                                }
                                if (fieldInfo.getHeadStyleProperty() != null) {
                                    BeanUtil.copyProperties(fieldInfo.getHeadStyleProperty(), styleProperty);
                                }
                                if (fieldInfo.getColumnWidthProperty() != null) {
                                    styleProperty.setWidth(fieldInfo.getColumnWidthProperty().getWidth());
                                }
                            });
                    setCellStyle(sheet, cell, styleProperty);
                } else {
                    setCellStyle(sheet, cell, headMap.get(i));
                }
            }
        }
    }

    private void setCellStyle(Sheet sheet, Cell cell, Head head) {
        CellStyleProperty cellStyleProperty = CellStyleProperty.getDefaultHeadProperty();
        if (head.getHeadStyleProperty() != null) {
            BeanUtil.copyProperties(head.getHeadStyleProperty(), cellStyleProperty);
            if (enabledNoBorder) {
                cellStyleProperty.setBorderBottom(BorderStyle.NONE);
                cellStyleProperty.setBorderTop(BorderStyle.NONE);
                cellStyleProperty.setBorderLeft(BorderStyle.NONE);
                cellStyleProperty.setBorderRight(BorderStyle.NONE);
            }
        }
        if (head.getColumnWidthProperty() != null) {
            cellStyleProperty.setWidth(head.getColumnWidthProperty().getWidth());
        }
        if (head.getHeadFontProperty() != null) {
            cellStyleProperty.setFontProperty(head.getHeadFontProperty());
        }
        setCellStyle(sheet, cell, cellStyleProperty);
    }

    private void setCellStyle(Sheet sheet, Cell cell, CellStyleProperty cellStyleProperty) {
        Workbook workbook = sheet.getWorkbook();
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.cloneStyleFrom(cell.getCellStyle());
        Font font = workbook.createFont();
        cellStyleProperty.setCellStyle(font, cellStyle);
        cell.setCellStyle(cellStyle);

        if (isAutoColumnWidth) {
            Short fontSize = cellStyleProperty.getFontProperty() == null ? font.getFontHeightInPoints() :
                    (cellStyleProperty.getFontProperty().getFontHeightInPoints() == null ?
                            font.getFontHeightInPoints() : cellStyleProperty.getFontProperty().getFontHeightInPoints());
            sheet.setColumnWidth(cell.getColumnIndex(), columnWidth(cell.getStringCellValue(), fontSize));
        } else if (cellStyleProperty.getWidth() != null) {
            sheet.setColumnWidth(cell.getColumnIndex(), cellStyleProperty.getWidth());
        }
    }

    /**
     * 根据表头计算列宽
     */
    private int columnWidth(String value, int fontSize) {
        String[] strs = value.split("\n");
        int result = 0;
        for (String s : strs) {
            int chineseNum = StringUtils.chineseNum(s);
            int dataLength = (int) ((s.length() - chineseNum) + chineseNum * 1.5);
            int columnWidth = dataLength * 40 * fontSize;
            columnWidth = columnWidth > Constants.MAX_COLUMN_WIDTH ? Constants.MAX_COLUMN_WIDTH : columnWidth;
            if (columnWidth > result) {
                result = columnWidth;
            }
        }
        return result;
    }

    public HeadStyleHandler isAutoColumnWidth(boolean isAutoColumnWidth) {
        this.isAutoColumnWidth = isAutoColumnWidth;
        return this;
    }

    public HeadStyleHandler setEnabledNoBorder(boolean enabledNoBorder) {
        this.enabledNoBorder = enabledNoBorder;
        return this;
    }

}
