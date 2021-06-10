package com.jz.zeus.excel.write.handler;

import cn.hutool.core.bean.BeanUtil;
import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.lang.Assert;
import cn.hutool.core.text.CharSequenceUtil;
import cn.hutool.core.util.StrUtil;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder;
import com.alibaba.excel.write.property.ExcelWriteHeadProperty;
import com.jz.zeus.excel.FieldInfo;
import com.jz.zeus.excel.annotation.HeadColor;
import com.jz.zeus.excel.context.ExcelContext;
import com.jz.zeus.excel.util.ClassUtils;
import com.jz.zeus.excel.util.ExcelUtils;
import com.jz.zeus.excel.write.helper.WriteSheetHelper;
import com.jz.zeus.excel.write.property.CellStyleProperty;
import org.apache.poi.ss.usermodel.*;

import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.stream.Collectors;

/**
 * @Author JZ
 * @Date 2021/3/23 11:57
 */
public class HeadStyleHandler extends AbstractRowWriteHandler {

    private ExcelContext excelContext;

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

    private CellStyleProperty allHeadStyle;

    /**
     * key 行索引、value 该行单元格的样式
     */
    private Map<Integer, List<CellStyleProperty>> headCellStyleMap;

    private WriteSheetHelper writeSheetHelper;

    /**
     * 若使用class表示表头且指定样式时，优先使用自定义样式，若无则使用class指定样式，若没有指定任何样式则使用默认样式。
     * 若是用list自定义表头则赋予默认表头样式
     */
    public HeadStyleHandler(ExcelContext excelContext) {
        Assert.notNull(excelContext, "ExcelContext must not be null");
        this.excelContext = excelContext;
    }

    /**
     * 为所有设置表头同一样式
     */
    public HeadStyleHandler(ExcelContext excelContext, CellStyleProperty allHeadStyle) {
        Assert.notNull(excelContext, "ExcelContext must not be null");
        this.excelContext = excelContext;
        this.allHeadStyle = allHeadStyle;
    }

    public HeadStyleHandler(ExcelContext excelContext, List<CellStyleProperty> headCellStyles) {
        Assert.notNull(excelContext, "ExcelContext must not be null");
        this.excelContext = excelContext;
        if (CollUtil.isNotEmpty(headCellStyles)) {
            this.headCellStyleMap = headCellStyles.stream()
                    .filter(e -> Objects.nonNull(e.getRowIndex()))
                    .collect(Collectors.groupingBy(CellStyleProperty::getRowIndex));
        }
    }

    @Override
    public void afterSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {
        writeSheetHelper = new WriteSheetHelper(excelContext, writeSheetHolder, null);
    }

    @Override
    public void afterRowDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, Row row, Integer relativeRowIndex, Boolean isHead) {
        if (!Boolean.TRUE.equals(isHead)) {
            return;
        }
        ExcelWriteHeadProperty excelWriteHeadProperty = writeSheetHolder.getExcelWriteHeadProperty();
        Class headClass = excelWriteHeadProperty.getHeadClazz();
        Map<Integer, Head> headMap = excelWriteHeadProperty.getHeadMap();
        Sheet sheet = writeSheetHolder.getSheet();
        int rawColumnNum = headMap.size();
        int realColumnNum = rawColumnNum + (CollUtil.isEmpty(excelContext.getExtendHead()) ?
                0 : excelContext.getExtendHead().size());
        for (int i = 0; i < realColumnNum; i++) {
            Cell cell = row.getCell(i);
            if (cell == null) {
                continue;
            }
            if (allHeadStyle != null) {
                setCellStyle(sheet, cell, allHeadStyle);
                continue;
            } else if (CollUtil.isNotEmpty(headCellStyleMap)) {
                List<CellStyleProperty> rowHeadCellStyles = headCellStyleMap.get(row.getRowNum());
                Integer finalI = i;
                CellStyleProperty styleProperty = CollUtil.isEmpty(rowHeadCellStyles) ? null :
                        rowHeadCellStyles.stream().filter(e -> finalI.equals(getColumnIndex(e))).findFirst().orElse(null);
                if (styleProperty != null) {
                    setCellStyle(sheet, cell, styleProperty);
                    continue;
                }
            }
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
                            HeadColor headColor = fieldInfo.getHeadColor();
                            if (headColor != null) {
                                if (CharSequenceUtil.isNotBlank(headColor.cellFillForegroundColor())) {
                                    styleProperty.setCellFillForegroundColor(headColor.cellFillForegroundColor());
                                }
                                if (CharSequenceUtil.isNotBlank(headColor.cellFillBackgroundColor())) {
                                    styleProperty.setCellFillBackgroundColor(headColor.cellFillBackgroundColor());
                                }
                            }
                        });
                setCellStyle(sheet, cell, styleProperty);
            } else {
                setCellStyle(sheet, cell, headMap.get(i), headClass);
            }
        }
    }

    private void setCellStyle(Sheet sheet, Cell cell, Head head, Class headClass) {
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
        ClassUtils.getFieldInfoByFieldName(headClass, head.getFieldName())
                .ifPresent(fieldInfo -> {
                    cellStyleProperty.setHeadColor(fieldInfo.getHeadColor());
                });
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
            sheet.setColumnWidth(cell.getColumnIndex(), ExcelUtils.calColumnWidth(cell.getStringCellValue(), fontSize));
        } else if (cellStyleProperty.getWidth() != null) {
            sheet.setColumnWidth(cell.getColumnIndex(), cellStyleProperty.getWidth());
        }
    }

    private Integer getColumnIndex(CellStyleProperty styleProperty) {
        Integer columnIndex = styleProperty.getColumnIndex();
        if (columnIndex == null) {
            if (StrUtil.isNotBlank(styleProperty.getFieldName())) {
                columnIndex = writeSheetHelper.getFieldColumnIndex(styleProperty.getFieldName());
            } else if (StrUtil.isNotBlank(styleProperty.getHeadName())) {
                columnIndex = writeSheetHelper.getHeadColumnIndex(styleProperty.getHeadName());
            }
        }
        return columnIndex;
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
