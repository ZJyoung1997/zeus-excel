package com.jz.zeus.excel.read.listener;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.text.CharSequenceUtil;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.enums.CellDataTypeEnum;
import com.alibaba.excel.enums.HeadKindEnum;
import com.alibaba.excel.event.AbstractIgnoreExceptionReadListener;
import com.alibaba.excel.exception.ExcelDataConvertException;
import com.alibaba.excel.metadata.CellData;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.metadata.property.ExcelContentProperty;
import com.alibaba.excel.read.metadata.holder.ReadHolder;
import com.alibaba.excel.read.metadata.property.ExcelReadHeadProperty;
import com.alibaba.excel.util.ConverterUtils;
import com.jz.zeus.excel.CellErrorInfo;
import com.jz.zeus.excel.exception.DataConvertException;
import net.sf.cglib.beans.BeanMap;

import java.util.*;

/**
 * @Author JZ
 * @Date 2021/6/25 16:34
 */
public class ModelBuildEventListener extends AbstractIgnoreExceptionReadListener<Map<Integer, CellData>> {

    @Override
    public void invokeHead(Map<Integer, CellData> cellDataMap, AnalysisContext context) {}

    @Override
    public void invoke(Map<Integer, CellData> cellDataMap, AnalysisContext context) {
        List<CellErrorInfo> cellErrorInfos = new ArrayList<>(cellDataMap.size());
        ReadHolder currentReadHolder = context.currentReadHolder();
        if (HeadKindEnum.CLASS == currentReadHolder.excelReadHeadProperty().getHeadKind()) {
            context.readRowHolder()
                    .setCurrentRowAnalysisResult(buildUserModel(cellDataMap, currentReadHolder, context, cellErrorInfos));
        } else {
            context.readRowHolder()
                    .setCurrentRowAnalysisResult(buildStringList(cellDataMap, currentReadHolder, context, cellErrorInfos));
        }
        if (CollUtil.isNotEmpty(cellErrorInfos)) {
            throw new DataConvertException(cellErrorInfos);
        }
    }

    private Object buildStringList(Map<Integer, CellData> cellDataMap, ReadHolder currentReadHolder,
                                   AnalysisContext context, List<CellErrorInfo> cellErrorInfos) {
        int index = 0;
        if (context.readWorkbookHolder().getDefaultReturnMap()) {
            Map<Integer, String> map = new LinkedHashMap<Integer, String>(cellDataMap.size() * 4 / 3 + 1);
            for (Map.Entry<Integer, CellData> entry : cellDataMap.entrySet()) {
                Integer key = entry.getKey();
                CellData cellData = entry.getValue();
                while (index < key) {
                    map.put(index, null);
                    index++;
                }
                index++;
                if (cellData.getType() == CellDataTypeEnum.EMPTY) {
                    map.put(key, null);
                    continue;
                }
                Integer rowIndex = context.readRowHolder().getRowIndex();
                String value = null;
                try {
                    value = (String) ConverterUtils.convertToJavaObject(cellData, null, null, currentReadHolder.converterMap(),
                            currentReadHolder.globalConfiguration(), rowIndex, key);
                } catch (ExcelDataConvertException e) {
                    cellErrorInfos.add(createCellErrorInfo(null, rowIndex, key, e));
                }
                map.put(key, value);
            }
            int headSize = currentReadHolder.excelReadHeadProperty().getHeadMap().size();
            while (index < headSize) {
                map.put(index, null);
                index++;
            }
            return map;
        } else {
            // Compatible with the old code the old code returns a list
            List<String> list = new ArrayList<String>();
            for (Map.Entry<Integer, CellData> entry : cellDataMap.entrySet()) {
                Integer key = entry.getKey();
                CellData cellData = entry.getValue();
                while (index < key) {
                    list.add(null);
                    index++;
                }
                index++;
                if (cellData.getType() == CellDataTypeEnum.EMPTY) {
                    list.add(null);
                    continue;
                }
                Integer rowIndex = context.readRowHolder().getRowIndex();
                String value = null;
                try {
                    value =(String) ConverterUtils.convertToJavaObject(cellData, null, null, currentReadHolder.converterMap(),
                            currentReadHolder.globalConfiguration(), rowIndex, key);
                } catch (ExcelDataConvertException e) {
                    cellErrorInfos.add(createCellErrorInfo(null, rowIndex, key, e));
                }
                list.add(value);
            }
            int headSize = currentReadHolder.excelReadHeadProperty().getHeadMap().size();
            while (index < headSize) {
                list.add(null);
                index++;
            }
            return list;
        }
    }

    private Object buildUserModel(Map<Integer, CellData> cellDataMap, ReadHolder currentReadHolder,
                                  AnalysisContext context, List<CellErrorInfo> cellErrorInfos) {
        ExcelReadHeadProperty excelReadHeadProperty = currentReadHolder.excelReadHeadProperty();
        Object resultModel;
        try {
            resultModel = excelReadHeadProperty.getHeadClazz().newInstance();
        } catch (Exception e) {
            throw new ExcelDataConvertException(context.readRowHolder().getRowIndex(), 0,
                    new CellData(CellDataTypeEnum.EMPTY), null,
                    "Can not instance class: " + excelReadHeadProperty.getHeadClazz().getName(), e);
        }
        Map<Integer, Head> headMap = excelReadHeadProperty.getHeadMap();
        Map<String, Object> map = new HashMap<String, Object>(headMap.size() * 4 / 3 + 1);
        Map<Integer, ExcelContentProperty> contentPropertyMap = excelReadHeadProperty.getContentPropertyMap();
        for (Map.Entry<Integer, Head> entry : headMap.entrySet()) {
            Integer index = entry.getKey();
            if (!cellDataMap.containsKey(index)) {
                continue;
            }
            CellData cellData = cellDataMap.get(index);
            if (cellData.getType() == CellDataTypeEnum.EMPTY) {
                continue;
            }
            ExcelContentProperty excelContentProperty = contentPropertyMap.get(index);
            Integer rowIndex = context.readRowHolder().getRowIndex();
            Object value = null;
            try {
                // 对于转化过程抛出的异常不做处理，可以方便后续进行错误记录
                value = ConverterUtils.convertToJavaObject(cellData, excelContentProperty.getField(),
                        excelContentProperty, currentReadHolder.converterMap(), currentReadHolder.globalConfiguration(),
                        rowIndex, index);
            } catch (ExcelDataConvertException e) {
                cellErrorInfos.add(createCellErrorInfo(entry.getValue(), rowIndex, index, e));
            }
            if (value != null) {
                map.put(excelContentProperty.getField().getName(), value);
            }
        }
        BeanMap.create(resultModel).putAll(map);
        return resultModel;
    }

    private CellErrorInfo createCellErrorInfo(Head head, Integer rowIndex, Integer columnIndex, ExcelDataConvertException e) {
        String errorMsg = "数据格式错误";
        if (e.getCause() instanceof DataConvertException) {
            errorMsg = e.getCause().getMessage();
        }
        if (head != null && CharSequenceUtil.isNotBlank(head.getFieldName())) {
            return CellErrorInfo.buildByField(rowIndex, head.getFieldName(), errorMsg);
        } else {
            return CellErrorInfo.buildByColumnIndex(rowIndex, columnIndex, errorMsg);
        }
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {}

}
