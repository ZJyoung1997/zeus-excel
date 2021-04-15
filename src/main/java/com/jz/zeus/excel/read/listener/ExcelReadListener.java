package com.jz.zeus.excel.read.listener;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.util.StrUtil;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.exception.ExcelDataConvertException;
import com.alibaba.excel.metadata.Cell;
import com.alibaba.excel.metadata.CellData;
import com.alibaba.excel.metadata.CellExtra;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.read.listener.ReadListener;
import com.alibaba.excel.read.metadata.holder.ReadRowHolder;
import com.alibaba.excel.read.metadata.property.ExcelReadHeadProperty;
import com.jz.zeus.excel.CellErrorInfo;
import com.jz.zeus.excel.FieldInfo;
import com.jz.zeus.excel.exception.DataConvertException;
import com.jz.zeus.excel.util.ClassUtils;
import com.jz.zeus.excel.util.UnsafeFieldAccessor;
import com.jz.zeus.excel.util.ValidatorUtils;
import lombok.Getter;
import lombok.Setter;

import java.lang.reflect.Field;
import java.util.*;
import java.util.stream.Collectors;

/**
 * @Author JZ
 * @Date 2021/3/22 14:31
 */
public abstract class ExcelReadListener<T> implements ReadListener<T> {

    /**
     * 读取Excel完毕后，Excel的真实表头行数
     */
    @Getter
    private Integer readAfterHeadRowNum;

    /**
     * 开启hibernate注解校验
     * true 开启、false 关闭
     */
    private boolean enabledAnnotationValidation = true;

    /**
     * 表头是否有误
     * false 正常、true 表头错误
     */
    @Getter
    private boolean headError = false;

    /**
     * 批量处理数据数量
     */
    @Setter
    private int batchHandleNum = 2000;

    /**
     * 表头错误信息
     */
    @Setter
    @Getter
    private String headErrorMsg;

    /**
     * 数据的错误信息（不包含表头错误信息）
     * key 为行索引，value 为该行中单元格的错误信息
     */
    @Getter
    private Map<Integer, List<CellErrorInfo>> errorInfoMap;

    /**
     * T 中字段名与列索引映射
     * key T 中字段名、value 列索引
     */
    private Map<String, Integer> fieldColumnIndexMap = new HashMap<>();

    /**
     * 表头与列索引映射
     * key 表头、value 表头对应列索引
     */
    private Map<String, Integer> headNameIndexMap = new HashMap<>();

    /**
     * 该字段中保存的是class中未定义的表头，
     * key 为列索引、value 为表头
     */
    private Map<Integer, String> dynamicColumnMap = new HashMap<>();

    /**
     * Excel 中读取到的数据
     * key 行索引，value 读取到的数据
     */
    private Map<Integer, T> dataMap = new HashMap<>();

    public ExcelReadListener() {}

    public ExcelReadListener(Boolean enabledAnnotationValidation) {
        this(enabledAnnotationValidation, null);
    }

    public ExcelReadListener(Integer batchHandleNum) {
        this(null, batchHandleNum);
    }

    public ExcelReadListener(Boolean enabledAnnotationValidation, Integer batchHandleNum) {
        if (enabledAnnotationValidation != null) {
            this.enabledAnnotationValidation = enabledAnnotationValidation;
        }
        if (batchHandleNum != null) {
            this.batchHandleNum = batchHandleNum;
        }
    }

    /**
     * 所有数据处理完毕之后触发的操作，如果想要所有数据校验完毕后在对数据进行操作
     * 可以对 {@link #dataHandle} 进行一个空的实现
     */
    protected abstract void doAfterAllDataHandle(AnalysisContext analysisContext);

    /**
     * 对读取到的数据进行处理
     * @param dataMap                   key 为行索引，value 为 Excel中该行数据
     * @param analysisContext
     */
    protected abstract void dataHandle(Map<Integer, T> dataMap, AnalysisContext analysisContext);

    /**
     * 校验读取到的数据
     * @param dataMap                   key 为行索引，value 为 Excel中该行数据
     * @param analysisContext
     */
    protected abstract void verify(Map<Integer, T> dataMap, AnalysisContext analysisContext);

    /**
     * 校验表头是否正常
     * 可调用 ConverterUtils.convertToStringMap(headMap, context) 方法将表头转化为对应map
     * @return 正常 true、异常 false
     */
    protected abstract void headCheck(Map<Integer, CellData> headMap, AnalysisContext context);

    @Override
    public void invoke(T data, AnalysisContext analysisContext) {
        ReadRowHolder readRowHolder = analysisContext.readRowHolder();
        annotationValidation(data, readRowHolder);
        Integer rowIndex = readRowHolder.getRowIndex();
        setDynamicColumnData(data, readRowHolder);
        dataMap.put(rowIndex, data);
        if (dataMap.size() >= batchHandleNum) {
            verify(dataMap, analysisContext);
            dataHandle(dataMap, analysisContext);
            dataMap.clear();
        }
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {
        if (dataMap.size() > 0) {
            verify(dataMap, analysisContext);
            dataHandle(dataMap, analysisContext);
            dataMap.clear();
        }
        doAfterAllDataHandle(analysisContext);
    }

    @Override
    public void invokeHead(Map<Integer, CellData> headMap, AnalysisContext analysisContext) {
        headError = false;
        headCheck(headMap, analysisContext);
        if (StrUtil.isNotBlank(this.headErrorMsg)) {
            headError = true;
        }
        initHeadCache(headMap,analysisContext.readSheetHolder().excelReadHeadProperty());
    }

    /**
     * 当表头有误时，停止Excel的读取
     * @param context
     * @return
     */
    @Override
    public boolean hasNext(AnalysisContext context) {
        if (headError) {
            return false;
        }
        return true;
    }

    @Override
    public void onException(Exception e, AnalysisContext analysisContext) {
        if (e instanceof ExcelDataConvertException) {
            String errorMsg = "数据类型错误";
            if (e.getCause() instanceof DataConvertException) {
                errorMsg = ((DataConvertException) e.getCause()).getErrorMsg();
            }
            ExcelDataConvertException dataConvertException = (ExcelDataConvertException) e;
            addErrorInfo(dataConvertException.getRowIndex(), dataConvertException.getColumnIndex(), errorMsg);
        }
    }

    @Override
    public void extra(CellExtra cellExtra, AnalysisContext analysisContext) {}

    /**
     * 获取动态数据
     */
    private void setDynamicColumnData(T data, ReadRowHolder readRowHolder) {
        if (CollUtil.isEmpty(dynamicColumnMap)) {
            return;
        }
        Map<Integer, Cell> cellDataMap = readRowHolder.getCellMap();
        if (CollUtil.isEmpty(cellDataMap)) {
            return;
        }
        Map<String, String> dynamicData = new HashMap<>(cellDataMap.size());
        cellDataMap.forEach((columnIndex, cell) -> {
            if (dynamicColumnMap.keySet().contains(columnIndex)) {
                dynamicData.put(dynamicColumnMap.get(columnIndex), cell.toString());
            }
        });
        ClassUtils.getClassFieldInfo(data.getClass()).stream()
                .filter(FieldInfo::isExtendColumn)
                .findFirst().ifPresent(fieldInfo -> {
            Field field = fieldInfo.getField();
            UnsafeFieldAccessor accessor = new UnsafeFieldAccessor(field);
            if (Map.class == field.getType()) {
                accessor.setObject(data, dynamicData);
            }
        });
    }

    protected void addErrorInfoByField(Integer rowIndex, String fieldName, String... errorMessages) {
        Integer columnIndex;
        if (rowIndex == null || (columnIndex = fieldColumnIndexMap.get(fieldName)) == null) {
            return;
        }
        addErrorInfo(rowIndex, columnIndex, errorMessages);
    }

    protected void addErrorInfoByHead(Integer rowIndex, String headName, String... errorMessages) {
        Integer columnIndex;
        if (rowIndex == null || (columnIndex = headNameIndexMap.get(headName)) == null) {
            return;
        }
        addErrorInfo(rowIndex, columnIndex, errorMessages);
    }

    protected void addErrorInfo(Integer rowIndex, Integer columnIndex, String... errorMessage) {
        if (rowIndex == null || columnIndex == null) {
            return;
        }
        if (this.errorInfoMap == null) {
            this.errorInfoMap = new HashMap<>();
        }
        List<CellErrorInfo> rowErrorInfoList = errorInfoMap.get(rowIndex);
        if (rowErrorInfoList == null) {
            rowErrorInfoList = new ArrayList<>();
            this.errorInfoMap.put(rowIndex, rowErrorInfoList);
        }
        rowErrorInfoList.add(new CellErrorInfo(rowIndex, columnIndex, errorMessage));
    }

    protected void addErrorInfo(List<CellErrorInfo> cellErrorInfos) {
        if (CollUtil.isEmpty(cellErrorInfos)) {
            return;
        }
        Map<Integer, List<CellErrorInfo>> tempMap = cellErrorInfos.stream()
                .collect(Collectors.groupingBy(CellErrorInfo::getRowIndex));
        if (this.errorInfoMap == null) {
            this.errorInfoMap = tempMap;
        } else {
            tempMap.forEach((rowIndex, errorInfos) -> {
                List<CellErrorInfo> cellErrorInfoList = this.errorInfoMap.get(rowIndex);
                if (cellErrorInfoList == null) {
                    this.errorInfoMap.put(rowIndex, errorInfos);
                } else {
                    cellErrorInfoList.addAll(errorInfos);
                }
            });
        }
    }

    /**
     * 对数据进行 hibernate 注解校验
     */
    protected void annotationValidation(T data, ReadRowHolder readRowHolder) {
        if (!enabledAnnotationValidation) {
            return;
        }
        Map<String, List<String>> errorMessageMap = ValidatorUtils.validate(data);
        if (CollUtil.isEmpty(errorMessageMap)) {
            return;
        }
        List<CellErrorInfo> errorMesInfoList = new ArrayList<>();
        errorMessageMap.forEach((fieldName, errorMessages) -> {
            errorMesInfoList.add(new CellErrorInfo(readRowHolder.getRowIndex(), fieldColumnIndexMap.get(fieldName))
                                .addErrorMsg(errorMessages));

        });
        addErrorInfo(errorMesInfoList);
    }

    /**
     * 判断是否存在表头错误
     * @return true 存在表头错误、false 不存在表头错误
     */
    public boolean hasHeadError() {
        return this.headError;
    }

    /**
     * 判断是否存在数据错误
     * @return true 存在数据错误、false 不存在数据错误
     */
    public boolean hasDataError() {
        return CollUtil.isNotEmpty(this.errorInfoMap);
    }

    /**
     * 判断当前时刻，行索引为 rowIndex 的行是否存在数据错误
     * @param rowIndex  行索引
     */
    public boolean hasDataErrorOnRow(Integer rowIndex) {
        if (!hasDataError()) {
            return true;
        }
        return CollUtil.isNotEmpty(this.errorInfoMap.get(rowIndex));
    }

    public List<CellErrorInfo> getErrorInfoList() {
        if (CollUtil.isEmpty(this.errorInfoMap)) {
            return Collections.emptyList();
        }
        List<CellErrorInfo> cellErrorInfoList = new ArrayList<>();
        this.errorInfoMap.values().forEach(errorInfos -> {
            cellErrorInfoList.addAll(errorInfos);
        });
        return cellErrorInfoList;
    }

    protected Integer getHeadIndex(String headName) {
        return headNameIndexMap.get(headName);
    }

    protected void initHeadCache(Map<Integer, CellData> headMap, ExcelReadHeadProperty excelHeadPropertyData) {
        fieldColumnIndexMap.clear();
        headNameIndexMap.clear();
        dynamicColumnMap.clear();

        readAfterHeadRowNum = excelHeadPropertyData.getHeadRowNumber();
        Map<Integer, Head> headMapData = excelHeadPropertyData.getHeadMap();
        Set<Integer> columnIndexSet = new HashSet<>();
        if (CollUtil.isNotEmpty(headMapData)) {
            headMapData.values().forEach(head -> {
                Integer columnIndex = head.getColumnIndex();
                columnIndexSet.add(columnIndex);
                if (StrUtil.isNotBlank(head.getFieldName())) {
                    fieldColumnIndexMap.put(head.getFieldName(), columnIndex);
                }
                head.getHeadNameList().forEach(headName -> {
                    headNameIndexMap.put(headName, columnIndex);
                });
            });
        }

        if (CollUtil.isNotEmpty(headMap)) {
            headMap.forEach((columnIndex, cellData) -> {
                String headName = cellData.toString();
                if (!headNameIndexMap.containsKey(headName)) {
                    headNameIndexMap.put(headName, columnIndex);
                }
                if (!columnIndexSet.contains(columnIndex)) {
                    dynamicColumnMap.put(columnIndex, headName);
                }
            });
        }
    }

}
