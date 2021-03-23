package com.jz.zeus.excel.write.handler;

import com.alibaba.excel.metadata.CellData;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.write.handler.AbstractCellWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import com.jz.zeus.excel.util.StringUtils;
import org.apache.poi.ss.usermodel.*;

import java.util.List;

/**
 * @Author JZ
 * @Date 2021/3/23 11:57
 */
public class DefaultHeadStyleHandler extends AbstractCellWriteHandler {

    private static final int MAX_COLUMN_WIDTH = 255*256;

    private static final int DEFAULT_HEAD_FONT_SIZE = 18;

    private int fontSize;

    public DefaultHeadStyleHandler() {
        this(DEFAULT_HEAD_FONT_SIZE);
    }

    public DefaultHeadStyleHandler(Integer fontSize) {
        if (fontSize == null) {
            this.fontSize = DEFAULT_HEAD_FONT_SIZE;
        } else {
            this.fontSize = fontSize;
        }
    }


    @Override
    public void afterCellDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, List<CellData> cellDataList, Cell cell, Head head, Integer relativeRowIndex, Boolean isHead) {
        if (!Boolean.TRUE.equals(isHead)) {
            return;
        }
        Sheet sheet = writeSheetHolder.getSheet();

        sheet.setColumnWidth(cell.getColumnIndex(), columnWidth(cell.getStringCellValue()));

        Workbook workbook = sheet.getWorkbook();
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setFillPattern(FillPatternType.NO_FILL);
        cellStyle.setBorderLeft(BorderStyle.NONE);
        cellStyle.setBorderRight(BorderStyle.NONE);
        cellStyle.setBorderTop(BorderStyle.NONE);
        cellStyle.setBorderBottom(BorderStyle.NONE);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cell.setCellStyle(cellStyle);

        Font font = workbook.createFont();
        font.setFontHeightInPoints((short) fontSize);
        font.setBold(false);
        font.setFontName("黑体");
        cellStyle.setFont(font);
    }

    /**
     * 根据表头计算列宽
     */
    private int columnWidth(String value) {
        int chineseNum = StringUtils.chineseNum(value);
        int dataLength = (int) (value.length() - chineseNum + chineseNum * 1.4);
        int columnWidth = dataLength * 36 * fontSize;
        return columnWidth > MAX_COLUMN_WIDTH ? MAX_COLUMN_WIDTH : columnWidth;
    }

}
