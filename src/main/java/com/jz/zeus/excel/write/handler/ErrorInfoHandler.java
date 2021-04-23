package com.jz.zeus.excel.write.handler;

import cn.hutool.core.collection.CollUtil;
import com.alibaba.excel.metadata.CellData;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.write.handler.AbstractCellWriteHandler;
import com.alibaba.excel.write.handler.SheetWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder;
import com.jz.zeus.excel.CellErrorInfo;
import com.jz.zeus.excel.context.ExcelContext;
import com.jz.zeus.excel.util.ExcelUtils;
import com.jz.zeus.excel.write.helper.WriteSheetHelper;
import lombok.Setter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * @Author JZ
 * @Date 2021/3/26 10:24
 */
public class ErrorInfoHandler extends AbstractCellWriteHandler implements SheetWriteHandler {

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
            this.rowErrorInfoMap = errorInfoList.stream()
                    .collect(Collectors.groupingBy(CellErrorInfo::getRowIndex));
        }
    }

    @Override
    public void afterCellDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, List<CellData> cellDataList, Cell cell, Head head, Integer relativeRowIndex, Boolean isHead) {
        if (writeSheetHelper == null) {
            return;
        }
        int rowIndex = cell.getRowIndex();
        if (CollUtil.isEmpty(rowErrorInfoMap.get(rowIndex))) {
            return;
        }
        Sheet sheet = writeSheetHolder.getCachedSheet();
        Iterator<CellErrorInfo> iterator = rowErrorInfoMap.get(rowIndex).iterator();
        while (iterator.hasNext()) {
            CellErrorInfo errorInfo = iterator.next();
            Integer columnIndex = getColumnIndex(errorInfo);
            if (columnIndex == null) {
                continue;
            }
            if (columnIndex == cell.getColumnIndex()) {
                ExcelUtils.setCommentErrorInfo(sheet, rowIndex, columnIndex, errorInfo.getErrorMsgs().toArray(new String[0]));
                iterator.remove();
                return;
            }
        }
    }

    @Override
    public void afterSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {
        if (CollUtil.isEmpty(rowErrorInfoMap)) {
            return;
        }
        writeSheetHelper = new WriteSheetHelper(excelContext, writeSheetHolder, headRowNum);
    }

    @Override
    public void beforeSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {}

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
