package com.jz.zeus.excel.test.converter;

import com.alibaba.excel.converters.Converter;
import com.alibaba.excel.enums.CellDataTypeEnum;
import com.alibaba.excel.metadata.CellData;
import com.alibaba.excel.metadata.GlobalConfiguration;
import com.alibaba.excel.metadata.property.ExcelContentProperty;
import com.jz.zeus.excel.exception.DataConvertException;

import java.math.BigDecimal;

/**
 * @Author JZ
 * @Date 2021/4/9 11:48
 */
public class LongConverter implements Converter<Long> {

    @Override
    public Class supportJavaTypeKey() {
        return Long.class;
    }

    @Override
    public CellDataTypeEnum supportExcelTypeKey() {
        return CellDataTypeEnum.NUMBER;
    }

    @Override
    public Long convertToJavaData(CellData cellData, ExcelContentProperty contentProperty, GlobalConfiguration globalConfiguration) throws Exception {
        try {
            return cellData.getNumberValue().longValue();
        } catch (Exception e) {
            throw new DataConvertException("数据类型应为 Long");
        }
    }

    @Override
    public CellData convertToExcelData(Long value, ExcelContentProperty contentProperty, GlobalConfiguration globalConfiguration) throws Exception {
        return new CellData(BigDecimal.valueOf(value));
    }
}
