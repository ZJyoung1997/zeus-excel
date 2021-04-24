package com.jz.zeus.excel.util;

import cn.hutool.core.util.ArrayUtil;
import lombok.experimental.UtilityClass;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

/**
 * @Author JZ
 * @Date 2021/3/26 11:52
 */
@UtilityClass
public class ExcelUtils {

    public void setCommentErrorInfo(Sheet sheet, Integer rowIndex, Integer columnIndex, String... errorMessages) {
        setCommentErrorInfo(sheet, rowIndex, columnIndex, "- ", "", errorMessages);
    }

    /**
     * 给sheet的指定单元格设置错误信息，错误信息将显示在批注里，且单元格背景色为红色
     * @param sheet                 单元格 所在sheet
     * @param rowIndex              单元格 的行索引
     * @param columnIndex           单元格 的列索引
     * @param errorMsgPrefix        错误信息前缀
     * @param errorMsgSuffix        错误信息后缀
     * @param errorMessages         错误信息
     */
    public void setCommentErrorInfo(Sheet sheet, Integer rowIndex, Integer columnIndex,
                                    String errorMsgPrefix, String errorMsgSuffix, String... errorMessages) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            return;
        }
        Cell cell = row.getCell(columnIndex);
        if (cell == null) {
            return;
        }
        CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
        cellStyle.cloneStyleFrom(cell.getCellStyle());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellStyle.setFillForegroundColor(IndexedColors.RED.index);
        cell.setCellStyle(cellStyle);

        Drawing<?> drawing = sheet.createDrawingPatriarch();
        Comment comment = drawing.createCellComment(new XSSFClientAnchor(0, 0, 0, 0, columnIndex, rowIndex, columnIndex+2, rowIndex+2));
        comment.setString(new XSSFRichTextString(ArrayUtil.join(errorMessages, "\n", errorMsgPrefix, errorMsgSuffix)));
        cell.setCellComment(comment);
    }

}
