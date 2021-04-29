package com.jz.zeus.excel.test;

import cn.hutool.core.collection.ListUtil;
import cn.hutool.core.io.IoUtil;
import cn.hutool.core.util.StrUtil;
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
        testCascade2007();
    }

    public static void testCascade2007() {
        // 查询所有的省名称
        List<String> provNameList = new ArrayList<String>();
        provNameList.add("安徽省");
        provNameList.add("浙江省");

        // 整理数据，放入一个Map中，mapkey存放父地点，value 存放该地点下的子区域
        Map<String, List<String>> siteMap = new HashMap<String, List<String>>();
        siteMap.put("浙江省", ListUtil.toList("杭州市", "宁波市"));
        siteMap.put("安徽省", ListUtil.toList("芜湖市", "滁州市"));
        siteMap.put("芜湖市", ListUtil.toList("戈江区", "三山区"));
        siteMap.put("滁州市", ListUtil.toList("来安县", "凤阳县"));

        // 创建一个excel
        Workbook book = new XSSFWorkbook();

        // 创建需要用户填写的数据页
        // 设计表头
        Sheet sheet1 = book.createSheet("sheet1");
        Row row0 = sheet1.createRow(0);
        row0.createCell(0).setCellValue("省");
        row0.createCell(1).setCellValue("市");
        row0.createCell(2).setCellValue("区");

        //创建一个专门用来存放地区信息的隐藏sheet页
        //因此也不能在现实页之前创建，否则无法隐藏。
        Sheet proviSheet = book.createSheet("proviSheet");
//        book.setSheetHidden(book.getSheetIndex(hideSheet), true);

        for(int i = 0; i < provNameList.size(); i ++){
            proviSheet.createRow(i).createCell(0).setCellValue(provNameList.get(i));
        }

        Sheet siteSheet = book.createSheet("siteSheet");
        siteMap.forEach((provi, siteList) -> {
            String s1 = null, s2 = null;
            for (int i = 0; i < siteList.size(); i++) {
                Cell cell = siteSheet.createRow(i).createCell(0);
                cell.setCellValue(siteList.get(i));
                if (i == 0) {
                    s1 = cell.getAddress().formatAsString();
                }
                if (i == siteList.size() - 1) {
                    s2 = cell.getAddress().formatAsString();
                }
            }

            if (book.getName(provi) != null) {
                return;
            }
            Name categoryName = book.createName();
            categoryName.setNameName(provi);
            String refersToFormula = StrUtil.strBuilder().append("siteSheet")
                    .append("!$").append(StrUtil.filter(s1, Character::isLetter))
                    .append('$').append(StrUtil.filter(s1, Character::isDigit))
                    .append(":$").append(StrUtil.filter(s2, Character::isLetter))
                    .append('$').append(StrUtil.filter(s2, Character::isDigit)).toString();
            categoryName.setRefersToFormula(refersToFormula);
        });

        XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper((XSSFSheet)sheet1);

        // 省规则
        DataValidationConstraint provConstraint = dvHelper.createExplicitListConstraint(provNameList.toArray(new String[]{}));
        CellRangeAddressList provRangeAddressList = new CellRangeAddressList(1, 30, 0, 0);
        DataValidation provinceDataValidation = dvHelper.createValidation(provConstraint, provRangeAddressList);
        provinceDataValidation.createErrorBox("error", "请选择正确的省份");
        provinceDataValidation.setShowErrorBox(true);
        provinceDataValidation.setSuppressDropDownArrow(true);
        sheet1.addValidationData(provinceDataValidation);


        // 市以规则，此处仅作一个示例
        // "INDIRECT($A$" + 2 + ")" 表示规则数据会从名称管理器中获取key与单元格 A2 值相同的数据，如果A2是浙江省，那么此处就是
        // 浙江省下的区域信息。
        DataValidationConstraint formula = dvHelper.createFormulaListConstraint("INDIRECT($A$" + 2 + ")");
        CellRangeAddressList rangeAddressList = new CellRangeAddressList(1,10,1,1);
        DataValidation cacse = dvHelper.createValidation(formula, rangeAddressList);
        cacse.createErrorBox("error", "请选择正确的市");
        sheet1.addValidationData(cacse);

        // 区规则
        formula = dvHelper.createFormulaListConstraint("INDIRECT($B$" + 2 + ")");
        rangeAddressList = new CellRangeAddressList(1,10,2,2);
        cacse = dvHelper.createValidation(formula, rangeAddressList);
        cacse.createErrorBox("error", "请选择正确的区");
        sheet1.addValidationData(cacse);


        FileOutputStream os = null;
        try {
            os = new FileOutputStream("C:\\Users\\User\\Desktop\\254.xlsx");
            book.write(os);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            IoUtil.close(os);
        }
    }

    /**
     *
     * @param offset 偏移量，如果给0，表示从A列开始，1，就是从B列
     * @param rowId 第几行
     * @param colCount 一共多少列
     * @return 如果给入参 1,1,10. 表示从B1-K1。最终返回 $B$1:$K$1
     *
     * @author denggonghai 2016年8月31日 下午5:17:49
     */
    public static String getRange(int offset, int rowId, int colCount) {
        char start = (char)('A' + offset);
        if (colCount <= 25) {
            char end = (char)(start + colCount - 1);
            return "$" + start + "$" + rowId + ":$" + end + "$" + rowId;
        } else {
            char endPrefix = 'A';
            char endSuffix = 'A';
            if ((colCount - 25) / 26 == 0 || colCount == 51) {// 26-51之间，包括边界（仅两次字母表计算）
                if ((colCount - 25) % 26 == 0) {// 边界值
                    endSuffix = (char)('A' + 25);
                } else {
                    endSuffix = (char)('A' + (colCount - 25) % 26 - 1);
                }
            } else {// 51以上
                if ((colCount - 25) % 26 == 0) {
                    endSuffix = (char)('A' + 25);
                    endPrefix = (char)(endPrefix + (colCount - 25) / 26 - 1);
                } else {
                    endSuffix = (char)('A' + (colCount - 25) % 26 - 1);
                    endPrefix = (char)(endPrefix + (colCount - 25) / 26);
                }
            }
            return "$" + start + "$" + rowId + ":$" + endPrefix + endSuffix + "$" + rowId;
        }
    }

}
