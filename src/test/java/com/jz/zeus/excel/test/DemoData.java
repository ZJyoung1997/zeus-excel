package com.jz.zeus.excel.test;

import com.alibaba.excel.annotation.ExcelIgnore;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ContentFontStyle;
import com.alibaba.excel.annotation.write.style.HeadFontStyle;
import com.alibaba.excel.annotation.write.style.HeadStyle;
import com.jz.zeus.excel.validation.IsLong;
import lombok.Data;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;

/**
 * @Author JZ
 * @Date 2021/2/23 18:00
 */
@Data
public class DemoData {

    @HeadStyle(fillPatternType = FillPatternType.NO_FILL,
            borderRight = BorderStyle.NONE, borderLeft = BorderStyle.NONE,
            borderBottom = BorderStyle.NONE, borderTop = BorderStyle.NONE)
    @HeadFontStyle(fontName = "微软雅黑", color = 10, bold = false)
    @ExcelProperty(value = {"ID"})
    private Long id;

    @IsLong
    @HeadStyle(fillPatternType = FillPatternType.NO_FILL,
            borderRight = BorderStyle.NONE, borderLeft = BorderStyle.NONE,
            borderBottom = BorderStyle.NONE, borderTop = BorderStyle.NONE)
    @HeadFontStyle(fontName = "微软雅黑", color = 10, bold = false)
    @ExcelProperty(value = "SRC")
    private String src;

    @HeadFontStyle(fontName = "微软雅黑", color = 10, bold = false)
    @ContentFontStyle(fontName = "微软雅黑", fontHeightInPoints = 14)
    @ExcelProperty(value = "DEST")
    private String dest;

    @ExcelProperty(value = "FUNC")
    private String func;

}
