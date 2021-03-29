package com.jz.zeus.excel.write.handler;

import com.alibaba.excel.metadata.CellData;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.write.handler.AbstractCellWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import com.jz.zeus.excel.util.StringUtils;
import com.jz.zeus.excel.write.property.StyleProperty;
import org.apache.poi.ss.usermodel.*;

import java.util.List;

/**
 * @Author JZ
 * @Date 2021/3/23 11:57
 */
public class DefaultHeadStyleHandler extends AbstractCellWriteHandler {

    private static final int MAX_COLUMN_WIDTH = 255*256;

    /**
     * 表头单元格样式，下标对应列索引
     */
    private List<StyleProperty> headStyles;

    public DefaultHeadStyleHandler() {
    }


    @Override
    public void afterCellDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, List<CellData> cellDataList, Cell cell, Head head, Integer relativeRowIndex, Boolean isHead) {
        if (!Boolean.TRUE.equals(isHead)) {
            return;
        }
        StyleProperty cellStyleProperty = StyleProperty.getDefaultHeadProperty();
        Sheet sheet = writeSheetHolder.getSheet();
        sheet.setColumnWidth(cell.getColumnIndex(),
                columnWidth(cell.getStringCellValue(), cellStyleProperty.getFontSize()));

        Workbook workbook = sheet.getWorkbook();
        CellStyle cellStyle = workbook.createCellStyle();

        cellStyleProperty.setCellStyle(workbook.createCellStyle());
        cell.setCellStyle(cellStyleProperty.setCellStyle(cellStyle));

        cellStyle.setFont(cellStyleProperty.setFontStyle(workbook.createFont()));
    }

    /**
     * 根据表头计算列宽
     */
    private int columnWidth(String value, int fontSize) {
        int chineseNum = StringUtils.chineseNum(value);
        int dataLength = (int) (value.length() - chineseNum + chineseNum * 1.4);
        int columnWidth = dataLength * 36 * fontSize;
        return columnWidth > MAX_COLUMN_WIDTH ? MAX_COLUMN_WIDTH : columnWidth;
    }

}
