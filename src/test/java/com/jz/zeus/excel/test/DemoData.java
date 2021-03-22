package com.jz.zeus.excel.test;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ContentFontStyle;
import com.alibaba.excel.annotation.write.style.ContentStyle;
import com.alibaba.excel.annotation.write.style.HeadFontStyle;
import com.alibaba.excel.annotation.write.style.HeadStyle;
import lombok.Data;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;

/**
 * @Author JZ
 * @Date 2021/2/23 18:00
 */
@Data
public class DemoData {

    @HeadStyle(fillPatternType = FillPatternType.NO_FILL,
            borderRight = BorderStyle.NONE, borderLeft = BorderStyle.NONE,
            borderBottom = BorderStyle.NONE, borderTop = BorderStyle.NONE)
    @HeadFontStyle(fontName = "黑体", color = 10, bold = false)
    @ContentStyle
    @ExcelProperty(value = "媒体CODE", index = 2)
    private String mateCode;

    @HeadStyle(fillPatternType = FillPatternType.NO_FILL,
            borderRight = BorderStyle.NONE, borderLeft = BorderStyle.NONE,
            borderBottom = BorderStyle.NONE, borderTop = BorderStyle.NONE)
    @ExcelProperty(value = "SRC", index = 1)
    private String src;

    @ContentFontStyle(fontName = "宋体", fontHeightInPoints = 14)
    @ExcelProperty(value = "DEST", index = 0)
    private String dest;

    @ExcelProperty(value = "FUNC", index = 3)
    private String func;

}
