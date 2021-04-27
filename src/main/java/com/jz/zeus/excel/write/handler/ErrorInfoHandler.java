package com.jz.zeus.excel.write.handler;

import cn.hutool.core.collection.CollUtil;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder;
import com.jz.zeus.excel.CellErrorInfo;
import com.jz.zeus.excel.context.ExcelContext;
import com.jz.zeus.excel.util.ExcelUtils;
import com.jz.zeus.excel.write.helper.WriteSheetHelper;
import lombok.Setter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * @Author JZ
 * @Date 2021/3/26 10:24
 */
public class ErrorInfoHandler extends AbstractRowWriteHandler {

    private ExcelContext excelContext;

    private WriteSheetHelper writeSheetHelper;

    private Integer headRowNum;

    @Setter
    private String commentRowPrefix = "- ";

    @Setter
    private String commentRowSuffix = "";

    private Map<Integer, List<CellErrorInfo>> rowErrorInfoMap;

    public ErrorInfoHandler(ExcelContext excelContext, List<CellErrorInfo> errorInfoList) {
        this(excelContext, null, errorInfoList);
    }

    /**
     * @param headRowNum     仅对用表头名作为列坐标时有影响，有误会导致加载不到对应表头，导致对应单元格无法添加错误信息
     * @param errorInfoList
     */
    public ErrorInfoHandler(ExcelContext excelContext, Integer headRowNum, List<CellErrorInfo> errorInfoList) {
        this.excelContext = excelContext;
        this.headRowNum = headRowNum;
        if (CollUtil.isNotEmpty(errorInfoList)) {
            rowErrorInfoMap = errorInfoList.stream()
                    .collect(Collectors.groupingBy(CellErrorInfo::getRowIndex));
        }
    }

    @Override
    public void afterRowDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, Row row, Integer relativeRowIndex, Boolean isHead) {
        int rowIndex = row.getRowNum();
        List<CellErrorInfo> currentRowErrorInfos = rowErrorInfoMap.get(rowIndex);
        if (writeSheetHelper == null || CollUtil.isEmpty(currentRowErrorInfos)) {
            return;
        }
        Sheet sheet = writeSheetHolder.getCachedSheet();
        for (CellErrorInfo errorInfo : currentRowErrorInfos) {
            Integer columnIndex = getColumnIndex(errorInfo);
            if (columnIndex != null) {
                ExcelUtils.setCommentErrorInfo(sheet, rowIndex, columnIndex, commentRowPrefix, commentRowSuffix,
                        errorInfo.getErrorMsgs().toArray(new String[0]));
            }
        }
        currentRowErrorInfos.clear();
    }

    @Override
    public void afterSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {
        if (CollUtil.isEmpty(rowErrorInfoMap)) {
            return;
        }
        writeSheetHelper = new WriteSheetHelper(excelContext, writeSheetHolder, headRowNum);

    }

    private Integer getColumnIndex(CellErrorInfo errorInfo) {
        Integer columnIndex = null;
        if ((columnIndex = errorInfo.getColumnIndex()) != null) {
            return columnIndex;
        } else if ((columnIndex = writeSheetHelper.getFieldColumnIndex(errorInfo.getFieldName())) != null) {
            return columnIndex;
        } else if ((columnIndex = writeSheetHelper.getHeadColumnIndex(errorInfo.getHeadName())) != null) {
            return columnIndex;
        }
        return columnIndex;
    }

}
