package com.jz.zeus.excel.util;

import cn.hutool.core.lang.Assert;
import cn.hutool.core.util.ArrayUtil;
import com.jz.zeus.excel.constant.Constants;
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

    /**
     * 根据字符长度和字号计算列宽
     */
    public int calColumnWidth(String cellStrValue, int fontSize) {
        String[] strs = cellStrValue.split("\n");
        int result = 0;
        for (String s : strs) {
            int chineseNum = StringUtils.chineseNum(s);
            int englishNum = s.length() - chineseNum;
            int columnWidth = (int) ((englishNum * 1.2 + chineseNum * 2) * 34 * fontSize);
            result = Math.max(result, columnWidth);
        }
        return result > Constants.MAX_COLUMN_WIDTH ? Constants.MAX_COLUMN_WIDTH : result;
    }

    public static int columnToIndex(String column) {
        if (!column.matches("[A-Z]+")) {
            try {
                throw new Exception("Invalid parameter");
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
        int index = 0;
        char[] chars = column.toUpperCase().toCharArray();
        for (int i = 0; i < chars.length; i++) {
            index += ((int) chars[i] - (int) 'A' + 1)
                    * (int) Math.pow(26, chars.length - i - 1);
        }
        return index;
    }

    /**
     * Excel列索引转字母
     * @param columnIndex      列索引,从 0 开始
     */
    public String columnIndexToStr(int columnIndex) {
        Assert.isTrue(columnIndex >= 0);
        StringBuilder column = new StringBuilder();
        do {
            if (column.length() > 0) {
                columnIndex--;
            }
            column.insert(0, ((char) (columnIndex % 26 + (int) 'A')));
            columnIndex = ((columnIndex - columnIndex % 26) / 26);
        } while (columnIndex > 0);
        return column.toString();
    }

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
