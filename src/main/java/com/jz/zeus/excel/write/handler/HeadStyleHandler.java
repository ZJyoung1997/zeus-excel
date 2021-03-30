package com.jz.zeus.excel.write.handler;

import cn.hutool.core.collection.CollUtil;
import com.alibaba.excel.metadata.CellData;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.write.handler.AbstractCellWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import com.jz.zeus.excel.constant.Constants;
import com.jz.zeus.excel.util.StringUtils;
import com.jz.zeus.excel.write.property.CellStyleProperty;
import org.apache.poi.ss.usermodel.*;

import java.util.List;

/**
 * @Author JZ
 * @Date 2021/3/23 11:57
 */
public class HeadStyleHandler extends AbstractCellWriteHandler {

    /**
     * 是否自适应列宽，默认开启
     */
    private boolean isAutoColumnWidth = true;

    /**
     * 表头单元格样式，外层list下标对应行索引，内层下标对应列索引
     */
    private List<List<CellStyleProperty>> multiRowHeadCellStyles;

    private CellStyleProperty allHeadStyle;

    /**
     * 为表头赋予一个默认的样式
     */
    public HeadStyleHandler() {
        this(CellStyleProperty.getDefaultHeadProperty());
    }

    /**
     * 为所有设置表头同一样式
     */
    public HeadStyleHandler(CellStyleProperty allHeadStyle) {
        this.allHeadStyle = allHeadStyle;
    }

    /**
     * 对表头的每个单元格按指定样式进行设置
     */
    public HeadStyleHandler(List<List<CellStyleProperty>> multiRowHeadCellStyles) {
        this.multiRowHeadCellStyles = multiRowHeadCellStyles;
    }


    @Override
    public void afterCellDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, List<CellData> cellDataList, Cell cell, Head head, Integer relativeRowIndex, Boolean isHead) {
        if (!Boolean.TRUE.equals(isHead)) {
            return;
        }
        Sheet sheet = writeSheetHolder.getSheet();

        if (allHeadStyle != null) {
            setCellStyle(sheet, cell, allHeadStyle);
        } else if (CollUtil.isNotEmpty(multiRowHeadCellStyles) && cell.getRowIndex() < multiRowHeadCellStyles.size()) {
            List<CellStyleProperty> cellStylePropertyList = multiRowHeadCellStyles.get(cell.getRowIndex());
            if (CollUtil.isNotEmpty(cellStylePropertyList) && cell.getColumnIndex() < cellStylePropertyList.size()) {
                setCellStyle(sheet, cell, cellStylePropertyList.get(cell.getColumnIndex()));
            }
        }

    }

    private void setCellStyle(Sheet sheet, Cell cell, CellStyleProperty cellStyleProperty) {
        Workbook workbook = sheet.getWorkbook();
        CellStyle cellStyle = workbook.createCellStyle();
        cell.setCellStyle(cellStyleProperty.setCellStyle(cellStyle));

        Font font = workbook.createFont();
        cellStyleProperty.setFontStyle(font);
        cellStyle.setFont(font);
        if (isAutoColumnWidth) {
            sheet.setColumnWidth(cell.getColumnIndex(),
                    columnWidth(cell.getStringCellValue(),
                            cellStyleProperty.getFontSize() == null ? font.getFontHeightInPoints() : cellStyleProperty.getFontSize())
            );
        } else if (cellStyleProperty.getWidth() != null) {
            sheet.setColumnWidth(cell.getColumnIndex(), cellStyleProperty.getWidth());
        }
    }

    /**
     * 根据表头计算列宽
     */
    private int columnWidth(String value, int fontSize) {
        int chineseNum = StringUtils.chineseNum(value);
        int dataLength = (int) ((value.length() - chineseNum) + chineseNum * 2);
        int columnWidth = dataLength * 36 * fontSize;
        return columnWidth > Constants.MAX_COLUMN_WIDTH ? Constants.MAX_COLUMN_WIDTH : columnWidth;
    }

    public HeadStyleHandler isAutoColumnWidth(boolean isAutoColumnWidth) {
        this.isAutoColumnWidth = isAutoColumnWidth;
        return this;
    }

}
