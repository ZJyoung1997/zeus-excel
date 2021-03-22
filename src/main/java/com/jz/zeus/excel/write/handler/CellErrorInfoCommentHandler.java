package com.jz.zeus.excel.write.handler;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.util.StrUtil;
import com.alibaba.excel.metadata.CellData;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.write.handler.AbstractCellWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import com.jz.zeus.excel.CellErrorInfo;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

import java.util.ArrayList;
import java.util.Collection;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * @Author JZ
 * @Date 2021/3/22 18:53
 */
public class CellErrorInfoCommentHandler extends AbstractCellWriteHandler {

    List<String> headList = new ArrayList<>();

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
        if (isHead) {
            headList.add(cell.getColumnIndex(), cell.getStringCellValue());
        }
        if (!rowErrorInfoMap.containsKey(cell.getRowIndex())) {
            return;
        }
        List<CellErrorInfo> cellErrorInfoList = rowErrorInfoMap.get(cell.getRowIndex());
        Collection<String> commentList = null;
        for (CellErrorInfo errorInfo : cellErrorInfoList) {
            if (errorInfo.getColumnIndex() != null && cell.getColumnIndex() == errorInfo.getColumnIndex().intValue()) {
                commentList = errorInfo.getErrorMsg();
                break;
            } else if (StrUtil.isNotBlank(errorInfo.getHeadName())
                    && errorInfo.getHeadName().equals(headList.get(cell.getColumnIndex()))) {
                commentList = errorInfo.getErrorMsg();
                break;
            }
        }
        if (CollUtil.isEmpty(commentList)) {
            return;
        }
        Drawing<?> drawing = writeSheetHolder.getSheet().createDrawingPatriarch();
        Comment comment = drawing.createCellComment(new XSSFClientAnchor(0, 0, 0, 0, cell.getColumnIndex(), cell.getRowIndex(), cell.getColumnIndex()+2, cell.getRowIndex()+2));
        comment.setString(new XSSFRichTextString(CollUtil.join(commentList, "\n")));
        cell.setCellComment(comment);
    }

}
