package com.jz.zeus.excel.write.handler;

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

import java.util.List;
import java.util.Map;

/**
 * @Author JZ
 * @Date 2021/3/22 18:53
 */
public class CommentWriteHandler extends AbstractCellWriteHandler {

    private Map<Integer, List<CellErrorInfo>> rowErrorInfoMap;


    @Override
    public void afterCellDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, List<CellData> cellDataList, Cell cell, Head head, Integer relativeRowIndex, Boolean isHead) {
        Drawing<?> drawing = writeSheetHolder.getSheet().createDrawingPatriarch();
        Comment comment = drawing.createCellComment(new XSSFClientAnchor());
        comment.setString(new XSSFRichTextString("批注"));
        cell.setCellComment(comment);
    }

}
