package com.jz.zeus.excel.test.data;

import com.alibaba.excel.annotation.ExcelIgnore;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.format.NumberFormat;
import com.alibaba.excel.annotation.write.style.ContentFontStyle;
import com.alibaba.excel.annotation.write.style.HeadFontStyle;
import com.alibaba.excel.annotation.write.style.HeadStyle;
import com.alibaba.excel.converters.longconverter.LongNumberConverter;
import com.jz.zeus.excel.annotation.ExtendColumn;
import com.jz.zeus.excel.annotation.HeadColor;
import com.jz.zeus.excel.annotation.ValidationData;
import com.jz.zeus.excel.test.converter.LongConverter;
import lombok.Data;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;

import java.math.BigDecimal;
import java.util.Map;

/**
 * @Author JZ
 * @Date 2021/2/23 18:00
 */
@Data
public class DemoData {

    @HeadColor(cellFillForegroundColor = "#8EA9DB")
    @HeadStyle(fillPatternType = FillPatternType.SOLID_FOREGROUND,
            fillForegroundColor = 31,
            borderRight = BorderStyle.NONE, borderLeft = BorderStyle.NONE,
            borderBottom = BorderStyle.NONE, borderTop = BorderStyle.NONE)
    @HeadFontStyle(fontName = "微软雅黑", color = 10, bold = false)
    @ExcelProperty(index = 0, value = "订单ID\n（不可修改）", converter = LongNumberConverter.class)
    private Long id;

//    @ExcelProperty(index = 1, value = {"SRC"}, converter = LongConverter.class)
    @ExcelProperty(index = 1, value = {"SRC"})
    private String src;

    @ExcelProperty(index = 2, value = "金额")
    private BigDecimal price;

//    @ContentFontStyle(fontName = "微软雅黑", fontHeightInPoints = 14)
//    @ExcelProperty(index = 3, value = "DEST", converter = LongConverter.class)
    @ExcelProperty(index = 3, value = "DEST")
    private String dest;

    @ValidationData(options = {"aa", "bb", "cc"}, errorTitle = "FUNC错误", errorMsg = "非法值")
    @ExcelProperty(index = 4, value = "FUNC")
    private String func;

    @ExcelProperty(index = 5, value = "省")
    private String provinces;

    @ExcelProperty(index = 6, value = "市")
    private String city;

    @ExcelProperty(index = 7, value = "镇")
    private String town;

    @ExcelIgnore
    @ExtendColumn
//    @HeadFontStyle(fontName = "微软雅黑", bold = false)
//    @ContentFontStyle(fontName = "微软雅黑", fontHeightInPoints = 14)
    private Map<String, String> extendColumnMap;


}
