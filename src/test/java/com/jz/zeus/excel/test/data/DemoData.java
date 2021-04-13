package com.jz.zeus.excel.test.data;

import com.alibaba.excel.annotation.ExcelIgnore;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.format.NumberFormat;
import com.alibaba.excel.annotation.write.style.ContentFontStyle;
import com.alibaba.excel.annotation.write.style.HeadFontStyle;
import com.alibaba.excel.annotation.write.style.HeadStyle;
import com.jz.zeus.excel.annotation.DynamicColumn;
import com.jz.zeus.excel.annotation.ValidationData;
import com.jz.zeus.excel.test.converter.LongConverter;
import com.jz.zeus.excel.validation.IsLong;
import lombok.Data;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.math.BigDecimal;
import java.util.Map;

/**
 * @Author JZ
 * @Date 2021/2/23 18:00
 */
@Data
public class DemoData {

    @HeadStyle(fillPatternType = FillPatternType.SOLID_FOREGROUND,
            fillForegroundColor = 31,
            borderRight = BorderStyle.NONE, borderLeft = BorderStyle.NONE,
            borderBottom = BorderStyle.NONE, borderTop = BorderStyle.NONE)
    @HeadFontStyle(fontName = "微软雅黑", color = 10, bold = false)
    @ExcelProperty(value = "订单ID\n（不可修改）", converter = LongConverter.class)
    private Long id;

    @IsLong
    @ExcelProperty(value = {"SRC"})
    private String src;

    @NumberFormat("#.#")
    @ExcelProperty(value = "金额")
    private Double price;

    @ContentFontStyle(fontName = "微软雅黑", fontHeightInPoints = 14)
    @ExcelProperty(value = "DEST")
    private String dest;

    @ValidationData(options = {"aa", "bb", "cc"})
    @ExcelProperty(value = "FUNC")
    private String func;

    @ExcelIgnore
    @DynamicColumn
    @HeadFontStyle(fontName = "微软雅黑", color = 10, bold = false)
    @ContentFontStyle(fontName = "微软雅黑", fontHeightInPoints = 14)
    private Map<String, String> dynamicColumnMap;

}
