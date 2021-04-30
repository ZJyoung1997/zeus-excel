package com.jz.zeus.excel.test;

import cn.hutool.core.collection.ListUtil;
import cn.hutool.core.io.IoUtil;
import cn.hutool.core.lang.Console;
import cn.hutool.core.util.StrUtil;
import com.jz.zeus.excel.util.ExcelUtils;
import lombok.SneakyThrows;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFDataValidationHelper;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.util.*;

/**
 * @Author JZ
 * @Date 2021/4/12 16:31
 */
public class Test {

    @SneakyThrows
    public static void main(String[] args) {
        Console.log(ExcelUtils.columnIndexToStr(43));
    }


}
