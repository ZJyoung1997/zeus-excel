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

        FontProperty fontProperty = new FontProperty();
        fontProperty.setFontHeightInPoints((short) 18);
        fontProperty.setBold(false);
        fontProperty.setFontName("微软雅黑");

        cellStyleProperty.setFontProperty(fontProperty);
        return cellStyleProperty;
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
        if (fontProperty.getCharset() != null) {
            font.setCharSet(fontProperty.getCharset());
        }
        if (fontProperty.getUnderline() != null) {
            font.setUnderline(fontProperty.getUnderline());
        }
        if (fontProperty.getTypeOffset() != null) {
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
        if (getDataFormat() != null) {
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
        if (getRotation() != null) {
            cellStyle.setRotation(getRotation());
        }
        if (getIndent() != null) {
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
        if (getLeftBorderColor() != null) {
            cellStyle.setLeftBorderColor(getLeftBorderColor());
        }
        if (getRightBorderColor() != null) {
            cellStyle.setRightBorderColor(getRightBorderColor());
        }
        if (getTopBorderColor() != null) {
            cellStyle.setTopBorderColor(getTopBorderColor());
        }
        if (getBottomBorderColor() != null) {
            cellStyle.setBottomBorderColor(getBottomBorderColor());
        }
        if (getFillPatternType() != null) {
            cellStyle.setFillPattern(getFillPatternType());
        }
        if (getFillBackgroundColor() != null) {
            cellStyle.setFillBackgroundColor(getFillBackgroundColor());
        }
        if (getFillForegroundColor() != null) {
            cellStyle.setFillForegroundColor(getFillForegroundColor());
        }
        if (getShrinkToFit() != null) {
            cellStyle.setShrinkToFit(getShrinkToFit());
        }
        return cellStyle;
    }

}
