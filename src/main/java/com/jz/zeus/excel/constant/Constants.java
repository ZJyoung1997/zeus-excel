package com.jz.zeus.excel.constant;

import org.apache.poi.xssf.usermodel.XSSFColor;

import java.awt.*;

/**
 * @Author JZ
 * @Date 2021/3/30 11:16
 */
public interface Constants {

    Integer MAX_COLUMN_WIDTH = 255*256;

    int colorIndex = new XSSFColor(Color.decode("#8EA9DB")).getIndexed();

}
