package com.jz.zeus.excel.write.property;

import cn.hutool.core.util.StrUtil;
import com.jz.zeus.excel.constant.Constants;
import lombok.Data;
import org.apache.poi.ss.usermodel.*;

/**
 * @author:JZ
 * @date:2021/3/29
 */
@Data
public class CellStyleProperty {

    /**
     * 宽度
     */
    private Integer width;

    /**
     * 高度
     */
    private Integer height;

    /**
     * 填充类型
     */
    private FillPatternType fillPattern;

    /**
     * 前景色
     */
    private Short fillForegroundColor;

    /**
     * 背景色
     */
    private Short fillBackgroundColor;

    /**
     * 左边框样式
     */
    private BorderStyle borderLeft;

    /**
     * 右边框样式
     */
    private BorderStyle borderRight;

    /**
     * 上边框样式
     */
    private BorderStyle borderTop;

    /**
     * 下边框样式
     */
    private BorderStyle borderBottom;

    /**
     * 水平对齐方式
     */
    private HorizontalAlignment horizontalAlignment;

    /**
     * 垂直对齐方式
     */
    private VerticalAlignment verticalAlignment;

    /**
     * 字体尺寸，同Excel中字体尺寸单位相同
     */
    private Integer fontSize;

    /**
     * 字体是否加粗，true 加粗、false 不加粗
     */
    private Boolean isBold;

    /**
     * 字体名称
     */
    private String fontName;

    public void setWidth(Integer width) {
        if (width != null && width > Constants.MAX_COLUMN_WIDTH) {
            this.width = Constants.MAX_COLUMN_WIDTH;
        } else {
            this.width = width;
        }
    }

    /**
     * 返回默认表头样式
     */
    public static CellStyleProperty getDefaultHeadProperty() {
        CellStyleProperty cellStyleProperty = new CellStyleProperty();
        cellStyleProperty.setFillPattern(FillPatternType.NO_FILL);
        cellStyleProperty.setBorderLeft(BorderStyle.NONE);
        cellStyleProperty.setBorderRight(BorderStyle.NONE);
        cellStyleProperty.setBorderTop(BorderStyle.NONE);
        cellStyleProperty.setBorderBottom(BorderStyle.NONE);
        cellStyleProperty.setHorizontalAlignment(HorizontalAlignment.CENTER);
        cellStyleProperty.setVerticalAlignment(VerticalAlignment.CENTER);

        cellStyleProperty.setFontSize(18);
        cellStyleProperty.setIsBold(false);
        cellStyleProperty.setFontName("黑体");
        return cellStyleProperty;
    }

    public void setFontStyle(Font font) {
        if (fontSize != null) {
            font.setFontHeightInPoints(Short.parseShort(fontSize.toString()));
        }
        if (isBold != null) {
            font.setBold(isBold);
        }
        if (StrUtil.isNotBlank(fontName)) {
            font.setFontName(fontName);
        }
        return;
    }

    public CellStyle setCellStyle(CellStyle cellStyle) {
        if (fillPattern != null) {
            cellStyle.setFillPattern(fillPattern);
        }
        if (fillForegroundColor != null) {
            cellStyle.setFillForegroundColor(fillForegroundColor);
        }
        if (fillBackgroundColor != null) {
            cellStyle.setFillBackgroundColor(fillBackgroundColor);
        }
        if (borderLeft != null) {
            cellStyle.setBorderLeft(borderLeft);
        }
        if (borderRight != null) {
            cellStyle.setBorderRight(borderRight);
        }
        if (borderTop != null) {
            cellStyle.setBorderTop(borderTop);
        }
        if (borderBottom != null) {
            cellStyle.setBorderBottom(borderBottom);
        }
        if (horizontalAlignment != null) {
            cellStyle.setAlignment(horizontalAlignment);
        }
        if (verticalAlignment != null) {
            cellStyle.setVerticalAlignment(verticalAlignment);
        }
        return cellStyle;
    }

}
