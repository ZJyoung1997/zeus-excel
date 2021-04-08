package com.jz.zeus.excel.write.property;

import cn.hutool.core.util.StrUtil;
import com.alibaba.excel.metadata.property.FontProperty;
import com.alibaba.excel.metadata.property.StyleProperty;
import com.jz.zeus.excel.constant.Constants;
import lombok.Getter;
import lombok.Setter;
import org.apache.poi.ss.usermodel.*;

/**
 * @author:JZ
 * @date:2021/3/29
 */
public class CellStyleProperty extends StyleProperty {

    public static final short DEFAULT_FONT_SIZE = 12;

    /**
     * 列宽
     */
    @Getter
    private Integer width;

    /**
     * 字体样式配置
     */
    @Getter
    @Setter
    private FontProperty fontProperty;

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
        cellStyleProperty.setFillPatternType(FillPatternType.NO_FILL);
        cellStyleProperty.setBorderLeft(BorderStyle.NONE);
        cellStyleProperty.setBorderRight(BorderStyle.NONE);
        cellStyleProperty.setBorderTop(BorderStyle.NONE);
        cellStyleProperty.setBorderBottom(BorderStyle.NONE);
        cellStyleProperty.setHorizontalAlignment(HorizontalAlignment.CENTER);
        cellStyleProperty.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyleProperty.setWrapped(true);
        cellStyleProperty.setFontProperty(getDefaultFontProperty());
        return cellStyleProperty;
    }

    public static FontProperty getDefaultFontProperty() {
        FontProperty fontProperty = new FontProperty();
        fontProperty.setFontHeightInPoints(DEFAULT_FONT_SIZE);
        fontProperty.setBold(true);
        fontProperty.setFontName("微软雅黑");
        return fontProperty;
    }

    public Font setFontStyle(Font font) {
        if (this.fontProperty == null) {
            return font;
        }
        if (fontProperty.getFontHeightInPoints() != null) {
            font.setFontHeightInPoints(fontProperty.getFontHeightInPoints());
        }
        if (fontProperty.getBold() != null) {
            font.setBold(fontProperty.getBold());
        }
        if (StrUtil.isNotBlank(fontProperty.getFontName())) {
            font.setFontName(fontProperty.getFontName());
        }
        if (fontProperty.getColor() != null) {
            font.setColor(fontProperty.getColor());
        }
        if (fontProperty.getItalic() != null) {
            font.setItalic(fontProperty.getItalic());
        }
        if (fontProperty.getCharset() != null && fontProperty.getCharset() != -1) {
            font.setCharSet(fontProperty.getCharset());
        }
        if (fontProperty.getUnderline() != null && fontProperty.getUnderline() != -1) {
            font.setUnderline(fontProperty.getUnderline());
        }
        if (fontProperty.getTypeOffset() != null && fontProperty.getTypeOffset() != -1) {
            font.setTypeOffset(fontProperty.getTypeOffset());
        }
        if (fontProperty.getStrikeout() != null) {
            font.setStrikeout(fontProperty.getStrikeout());
        }
        return font;
    }

    public void setCellStyle(Font font, CellStyle cellStyle) {
        setFontStyle(font);
        setCellStyle(cellStyle);
        cellStyle.setFont(font);
    }

    public CellStyle setCellStyle(CellStyle cellStyle) {
        if (getDataFormat() != null && getDataFormat() != -1) {
            cellStyle.setDataFormat(getDataFormat());
        }
        if (getHidden() != null) {
            cellStyle.setHidden(getHidden());
        }
        if (getLocked() != null) {
            cellStyle.setLocked(getLocked());
        }
        if (getQuotePrefix() != null) {
            cellStyle.setQuotePrefixed(getQuotePrefix());
        }
        if (getHorizontalAlignment() != null) {
            cellStyle.setAlignment(getHorizontalAlignment());
        }
        if (getWrapped() != null) {
            cellStyle.setWrapText(getWrapped());
        }
        if (getVerticalAlignment() != null) {
            cellStyle.setVerticalAlignment(getVerticalAlignment());
        }
        if (getRotation() != null && getRotation() != -1) {
            cellStyle.setRotation(getRotation());
        }
        if (getIndent() != null && getIndent() != -1) {
            cellStyle.setIndention(getIndent());
        }
        if (getBorderLeft() != null) {
            cellStyle.setBorderLeft(getBorderLeft());
        }
        if (getBorderRight() != null) {
            cellStyle.setBorderRight(getBorderRight());
        }
        if (getBorderTop() != null) {
            cellStyle.setBorderTop(getBorderTop());
        }
        if (getBorderBottom() != null) {
            cellStyle.setBorderBottom(getBorderBottom());
        }
        if (getLeftBorderColor() != null && getLeftBorderColor() != -1) {
            cellStyle.setLeftBorderColor(getLeftBorderColor());
        }
        if (getRightBorderColor() != null && getRightBorderColor() != -1) {
            cellStyle.setRightBorderColor(getRightBorderColor());
        }
        if (getTopBorderColor() != null && getTopBorderColor() != -1) {
            cellStyle.setTopBorderColor(getTopBorderColor());
        }
        if (getBottomBorderColor() != null && getBottomBorderColor() != -1) {
            cellStyle.setBottomBorderColor(getBottomBorderColor());
        }
        if (getFillPatternType() != null) {
            cellStyle.setFillPattern(getFillPatternType());
        }
        if (getFillBackgroundColor() != null && getFillBackgroundColor() != -1) {
            cellStyle.setFillBackgroundColor(getFillBackgroundColor());
        }
        if (getFillForegroundColor() != null && getFillForegroundColor() != -1) {
            cellStyle.setFillForegroundColor(getFillForegroundColor());
        }
        if (getShrinkToFit() != null) {
            cellStyle.setShrinkToFit(getShrinkToFit());
        }
        return cellStyle;
    }

}
