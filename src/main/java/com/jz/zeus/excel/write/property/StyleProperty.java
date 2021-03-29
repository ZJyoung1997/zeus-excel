package com.jz.zeus.excel.write.property;

import cn.hutool.core.util.StrUtil;
import lombok.Data;
import org.apache.poi.ss.usermodel.*;

/**
 * @author:JZ
 * @date:2021/3/29
 */
@Data
public class StyleProperty {

    /**
     * 填充类型
     */
    private FillPatternType fillPattern;

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

    /**
     * 返回默认表头样式
     */
    public static StyleProperty getDefaultHeadProperty() {
        StyleProperty styleProperty = new StyleProperty();
        styleProperty.setFillPattern(FillPatternType.NO_FILL);
        styleProperty.setBorderLeft(BorderStyle.NONE);
        styleProperty.setBorderRight(BorderStyle.NONE);
        styleProperty.setBorderTop(BorderStyle.NONE);
        styleProperty.setBorderBottom(BorderStyle.NONE);
        styleProperty.setHorizontalAlignment(HorizontalAlignment.CENTER);
        styleProperty.setVerticalAlignment(VerticalAlignment.CENTER);

        styleProperty.setFontSize(18);
        styleProperty.setIsBold(false);
        styleProperty.setFontName("黑体");
        return styleProperty;
    }

    public Font setFontStyle(Font font) {
        if (fontSize != null) {
            font.setFontHeightInPoints(Short.parseShort(fontSize.toString()));
        }
        if (isBold != null) {
            font.setBold(isBold);
        }
        if (StrUtil.isNotBlank(fontName)) {
            font.setFontName(fontName);
        }
        return font;
    }

    public CellStyle setCellStyle(CellStyle cellStyle) {
        if (fillPattern != null) {
            cellStyle.setFillPattern(fillPattern);
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
