package com.jz.zeus.excel.read.listener;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.map.MapUtil;
import cn.hutool.core.util.StrUtil;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.enums.HeadKindEnum;
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
import javafx.util.Pair;
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
    private Map<Integer, List<CellErrorInfo>> errorInfoMap = new HashMap<>();

    /**
     * 错误数据
     * key 行索引、value Excel中读取到的该行数据
     */
    private Map<Integer, T> errorDataMap = new HashMap<>();

    private Field extendField;

    /**
     * T 中字段名与列索引映射
     * key T 中字段名、value 列索引
     */
    private Map<String, Integer> fieldColumnIndexMap = new HashMap<>();

    /**
     * 该字段中保存的是class中未定义的表头，
     * key 为列索引、value 为表头
     */
    private Map<Integer, String> extendColumnIndexMap = new HashMap<>();

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
        setExtendColumnData(data, readRowHolder);
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
        initHeadCache(headMap, analysisContext.readSheetHolder().excelReadHeadProperty());
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
        } else {
            throw new RuntimeException(e);
        }
    }

    @Override
    public void extra(CellExtra cellExtra, AnalysisContext analysisContext) {}

    /**
     * 获取动态数据
     */
    private void setExtendColumnData(T data, ReadRowHolder readRowHolder) {
        if (extendField == null || CollUtil.isEmpty(extendColumnIndexMap)) {
            return;
        }
        Map<Integer, Cell> cellDataMap = readRowHolder.getCellMap();
        if (CollUtil.isEmpty(cellDataMap)) {
            return;
        }
        Map<String, String> extendData = new HashMap<>();
        cellDataMap.forEach((columnIndex, cell) -> {
            String extendHeadName = extendColumnIndexMap.get(columnIndex);
            if (StrUtil.isNotBlank(extendHeadName)) {
                extendData.put(extendHeadName, cell.toString());
            }
        });

        UnsafeFieldAccessor accessor = new UnsafeFieldAccessor(extendField);
        if (Map.class == extendField.getType()) {
            accessor.setObject(data, extendData);
        }
    }

    /**
     * 记录错误数据
     * @param rowIndex    行索引
     * @param data        错误数据
     */
    private void addErrorDataRecord(Integer rowIndex, T data) {
        errorDataMap.put(rowIndex, data);
    }

    /**
     * 根据行索引和扩展列表头添加错误信息，目前扩展列添加错误信息需通过该方法添加错误信息
     * @param rowIndex            行索引
     * @param extendHeadName      扩展列表头名
     * @param errorMessages       错误信息
     */
    protected void addExtendColumnErrorInfo(Integer rowIndex, String extendHeadName, String... errorMessages) {
        if (rowIndex == null || StrUtil.isBlank(extendHeadName)) {
            return;
        }
        Integer columnIndex = null;
        for (Map.Entry<Integer, String> entry : extendColumnIndexMap.entrySet()) {
            if (Objects.equals(extendHeadName, entry.getValue())) {
                columnIndex = entry.getKey();
                break;
            }
        }
        if (columnIndex == null) {
            return;
        }
        addErrorInfo(rowIndex, columnIndex, errorMessages);
    }

    /**
     * 根据行索引和属性名称添加错误信息，扩展列不适用于该方法
     * @param rowIndex             行索引
     * @param fieldName            class中属性名
     * @param errorMessages        错误信息
     */
    protected void addErrorInfo(Integer rowIndex, String fieldName, String... errorMessages) {
        Integer columnIndex;
        if (rowIndex == null || (columnIndex = fieldColumnIndexMap.get(fieldName)) == null) {
            return;
        }
        addErrorInfo(rowIndex, columnIndex, errorMessages);
    }

    /**
     * 根据行索引和列索引添加错误信息
     * @param rowIndex          行索引
     * @param columnIndex       列索引
     * @param errorMessage      错误信息
     */
    protected void addErrorInfo(Integer rowIndex, Integer columnIndex, String... errorMessage) {
        if (rowIndex == null || columnIndex == null) {
            return;
        }
        List<CellErrorInfo> rowErrorInfoList = errorInfoMap.get(rowIndex);
        if (rowErrorInfoList == null) {
            rowErrorInfoList = new ArrayList<>();
            this.errorInfoMap.put(rowIndex, rowErrorInfoList);
        }
        rowErrorInfoList.add(CellErrorInfo.buildByColumnIndex(rowIndex, columnIndex, errorMessage));
        addErrorDataRecord(rowIndex, dataMap.get(rowIndex));
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

    /**
     * 获取错误信息
     */
    public List<CellErrorInfo> getErrorInfoList() {
        if (CollUtil.isEmpty(this.errorInfoMap)) {
            return new ArrayList<>(0);
        }
        List<CellErrorInfo> cellErrorInfoList = new ArrayList<>();
        this.errorInfoMap.values().forEach(errorInfos -> {
            cellErrorInfoList.addAll(errorInfos);
        });
        return cellErrorInfoList;
    }

    /**
     * 获取错误数据及其错误信息，若要将其作为Excel的数据源创建新的Excel，并显示错误信息时，
     * 需修改错误信息的行索引
     * @return
     */
    public Pair<List<T>, List<CellErrorInfo>> getErrorRecord() {
        if (CollUtil.isEmpty(errorDataMap)) {
            return new Pair<>(null, null);
        }
        List<T> errorDatas = new ArrayList<>(errorDataMap.size());
        List<CellErrorInfo> newCellErrorInfos = new ArrayList<>();
        int rowIndex = readAfterHeadRowNum;
        for (Map.Entry<Integer, T> entry : errorDataMap.entrySet()) {
            errorDatas.add(entry.getValue());
            int finalRowIndex = rowIndex;
            CollUtil.addAll(newCellErrorInfos, errorInfoMap.get(entry.getKey()).stream()
                    .map(errorInfo -> errorInfo.clone().setRowIndex(finalRowIndex)).collect(Collectors.toList()));
            rowIndex++;
        }
        return new Pair<>(errorDatas, newCellErrorInfos);
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
            errorMesInfoList.add(CellErrorInfo.buildByColumnIndex(readRowHolder.getRowIndex(), fieldColumnIndexMap.get(fieldName), errorMessages));
        });
        errorMesInfoList.stream().collect(Collectors.groupingBy(CellErrorInfo::getRowIndex))
                .forEach((rowIndex, errorInfos) -> {
                    List<CellErrorInfo> infos = errorInfoMap.get(rowIndex);
                    if (infos == null) {
                        infos = new ArrayList<>(errorInfos.size());
                        errorInfoMap.put(rowIndex, infos);
                    }
                    infos.addAll(errorInfos);
                    addErrorDataRecord(rowIndex, data);
                });
    }

    protected void initHeadCache(Map<Integer, CellData> headMap, ExcelReadHeadProperty excelHeadPropertyData) {
        fieldColumnIndexMap.clear();
        extendColumnIndexMap.clear();
        extendField = null;

        readAfterHeadRowNum = excelHeadPropertyData.getHeadRowNumber();

        if (HeadKindEnum.CLASS == excelHeadPropertyData.getHeadKind()) {
            ClassUtils.getClassFieldInfo(excelHeadPropertyData.getHeadClazz()).stream()
                    .filter(FieldInfo::isExtendColumn)
                    .findFirst().ifPresent(fieldInfo -> extendField = fieldInfo.getField());
        }

        Map<Integer, Head> headMapData = excelHeadPropertyData.getHeadMap();
        Set<Integer> columnIndexSet = new HashSet<>();
        if (CollUtil.isNotEmpty(headMapData)) {
            headMapData.values().forEach(head -> {
                Integer columnIndex = head.getColumnIndex();
                columnIndexSet.add(columnIndex);
                if (StrUtil.isNotBlank(head.getFieldName())) {
                    fieldColumnIndexMap.put(head.getFieldName(), columnIndex);
                }
            });
        }

        if (CollUtil.isNotEmpty(headMap)) {
            headMap.forEach((columnIndex, cellData) -> {
                String headName = cellData.toString();
                if (!columnIndexSet.contains(columnIndex)) {
                    extendColumnIndexMap.put(columnIndex, headName);
                }
            });
        }
    }

}
