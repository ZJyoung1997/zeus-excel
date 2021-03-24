package com.jz.zeus.excel.write.handler;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.util.StrUtil;
import com.alibaba.excel.metadata.CellData;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import com.jz.zeus.excel.CellErrorInfo;
import lombok.Setter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

import java.util.Collection;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * @Author JZ
 * @Date 2021/3/22 18:53
 */
public class CellErrorInfoCommentHandler extends AbstractZeusCellWriteHandler {

    @Setter
    private String commentRowPrefix = "- ";

    @Setter
    private String commentRowSuffix = "";

    private Map<Integer, List<CellErrorInfo>> rowErrorInfoMap;

    public CellErrorInfoCommentHandler(Map<Integer, List<CellErrorInfo>> rowErrorInfoMap) {
        this.rowErrorInfoMap = rowErrorInfoMap;
    }

    public CellErrorInfoCommentHandler(List<CellErrorInfo> cellErrorInfoList) {
        if (CollUtil.isNotEmpty(cellErrorInfoList)) {
            this.rowErrorInfoMap = cellErrorInfoList.stream()
                    .collect(Collectors.groupingBy(CellErrorInfo::getRowIndex));
        }
    }

    @Override
    public void afterCellDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, List<CellData> cellDataList, Cell cell, Head head, Integer relativeRowIndex, Boolean isHead) {
        if (CollUtil.isEmpty(rowErrorInfoMap)) {
            return;
        }
        if (Boolean.TRUE.equals(isHead)) {
            loadExcelHead(writeSheetHolder, cell);
        }
        int rowIndex = cell.getRowIndex();
        int columnIndex = cell.getColumnIndex();
        if (!rowErrorInfoMap.containsKey(rowIndex)) {
            return;
        }
        List<CellErrorInfo> cellErrorInfoList = rowErrorInfoMap.get(rowIndex);
        Collection<String> commentList = null;
        for (CellErrorInfo errorInfo : cellErrorInfoList) {
            if (errorInfo.getColumnIndex() != null && columnIndex == errorInfo.getColumnIndex()) {
                commentList = errorInfo.getErrorMsgs();
                break;
            } else if (StrUtil.isNotBlank(errorInfo.getHeadName())
                    && errorInfo.getHeadName().equals(getHeadName(writeSheetHolder, columnIndex))) {
                commentList = errorInfo.getErrorMsgs();
                break;
            }
        }
        if (CollUtil.isEmpty(commentList)) {
            return;
        }
        Sheet sheet = writeSheetHolder.getSheet();
        CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellStyle.setFillForegroundColor(IndexedColors.RED.index);
        cell.setCellStyle(cellStyle);

        Drawing<?> drawing = sheet.createDrawingPatriarch();
        Comment comment = drawing.createCellComment(new XSSFClientAnchor(0, 0, 0, 0, columnIndex, rowIndex, columnIndex+2, rowIndex+2));
        comment.setString(new XSSFRichTextString(CollUtil.join(commentList, "\n", commentRowPrefix, commentRowSuffix)));
        cell.setCellComment(comment);
    }

}
