package com.jz.zeus.excel.test.data;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.HeadStyle;
import com.jz.zeus.excel.annotation.DropDownBox;
import lombok.Data;
import org.apache.poi.ss.usermodel.FillPatternType;

/**
 * @Author JZ
 * @Date 2021/4/8 15:12
 */
@Data
public class LineOrderPDBData {

    @ExcelProperty(value = "订单ID（不可修改）\n若留空则代表新建，否则代表编辑")
    private Long id;

    @HeadStyle(fillPatternType = FillPatternType.SOLID_FOREGROUND, fillForegroundColor = 31)
    @ExcelProperty(value = "订单名称（必填）")
    private String name;

    @HeadStyle(fillPatternType = FillPatternType.SOLID_FOREGROUND, fillForegroundColor = 31)
    @ExcelProperty(value = "媒体（不可修改）")
    private String mediaName;

    @DropDownBox(options = {"预加载", "非预加载"})
    @HeadStyle(fillPatternType = FillPatternType.SOLID_FOREGROUND, fillForegroundColor = 31)
    @ExcelProperty(value = "加载方式（不可修改）")
    private String loadType;

    @ExcelProperty(value = "市场（选填）")
    private String geos;

    @ExcelProperty(value = "终端（选填）")
    private String devices;

    @HeadStyle(fillPatternType = FillPatternType.SOLID_FOREGROUND, fillForegroundColor = 31)
    @ExcelProperty(value = "是否流量交换（不可修改）")
    private String isExchange;

    @HeadStyle(fillPatternType = FillPatternType.SOLID_FOREGROUND, fillForegroundColor = 31)
    @ExcelProperty(value = "流量交换类型\n（当流量交换时，必填，不可修改）")
    private String exchangeType;

    @ExcelProperty(value = "Deal ID（选填）")
    private String dealId;

    @ExcelProperty(value = "预算（选填）单位：￥\n（当交换订单时，该值无实际作用）")
    private String budget;

    @ExcelProperty(value = "投放量（选填）单位: 千次\n（当交换订单时，该值无实际作用）")
    private String bought;

    @HeadStyle(fillPatternType = FillPatternType.SOLID_FOREGROUND, fillForegroundColor = 31)
    @ExcelProperty(value = "是否固定推送比（非交换订单时，必填；交换订单时，不支持）")
    private String pushRateFixed;

    @HeadStyle(fillPatternType = FillPatternType.SOLID_FOREGROUND, fillForegroundColor = 31)
    @ExcelProperty(value = "约定推送比（非交换订单时，必填；交换订单时，不支持）\n示例：1.2")
    private String pushRate;

    @HeadStyle(fillPatternType = FillPatternType.SOLID_FOREGROUND, fillForegroundColor = 31)
    @ExcelProperty(value = "投放起止时间（非交换订单时，必填；交换订单时，不支持）\n示例：20200201 07:00 ~20200220 15:00")
    private String startAndEndTime;

    @ExcelProperty(value = "性别年龄定向（选填）\n示例：25-29男")
    private String tas;

    @ExcelProperty(value = "包含优选人群（选填）\n示例: 体育人群(1234)")
    private String includeTagsStr;

    @ExcelProperty(value = "排除优选人群（选填）\n示例: 体育人群(1234)")
    private String excludeTagsStr;

}
