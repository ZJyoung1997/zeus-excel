package com.jz.zeus.excel.write.handler;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.util.StrUtil;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.write.handler.AbstractSheetWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder;
import com.alibaba.excel.write.property.ExcelWriteHeadProperty;
import com.jz.zeus.excel.DynamicHead;

import java.util.List;
import java.util.Map;
import java.util.Objects;

/**
 * @author:JZ
 * @date:2021/4/13
 */
public class DynamicHeadHandler extends AbstractSheetWriteHandler {

    private List<DynamicHead> dynamicHeads;

    public DynamicHeadHandler(List<DynamicHead> dynamicHeads) {
        this.dynamicHeads = dynamicHeads;
    }

    @Override
    public void beforeSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {
        if (CollUtil.isEmpty(dynamicHeads)) {
            return;
        }
        ExcelWriteHeadProperty headProperty = writeSheetHolder.excelWriteHeadProperty();
        Map<Integer, Head> headMap = headProperty.getHeadMap();
        headMap.forEach((columnIndex, head) -> {
            dynamicHeads.stream()
                    .filter(dynamicHead -> Objects.equals(head.getColumnIndex(), dynamicHead.getColumnIndex())
                            || Objects.equals(head.getFieldName(), dynamicHead.getFieldName()))
                    .findFirst().ifPresent(dynamicHead -> {
                        List<String> headNames = head.getHeadNameList();
                        if (CollUtil.isEmpty(headNames)) {
                            return;
                        }
                        Integer rowIndex = dynamicHead.getRowIndex();
                        headNames.set(rowIndex, dynamicHead.getFinalHeadName(headNames.get(rowIndex)));
                    });
        });
    }

}
