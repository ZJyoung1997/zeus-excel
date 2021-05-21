package com.jz.zeus.excel.validator;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.text.CharSequenceUtil;
import cn.hutool.core.text.StrBuilder;
import cn.hutool.core.util.StrUtil;
import lombok.Getter;
import lombok.ToString;

import java.util.*;

/**
 * @Author JZ
 * @Date 2021/5/18 10:58
 */
@ToString
public class VerifyResult {

    @Getter
    private Map<String, List<String>> errorInfoMap = new HashMap<>();

    public void addErrorInfo(String fieldName, String errorMsg) {
        List<String> errorMsgs = errorInfoMap.get(fieldName);
        if (Objects.isNull(errorMsgs)) {
            errorMsgs = new ArrayList<>();
            errorInfoMap.put(fieldName, errorMsgs);
        }
        if (CharSequenceUtil.isNotBlank(errorMsg)) {
            errorMsgs.add(errorMsg);
        }
    }

    public void addErrorInfo(String fieldName, List<String> errorMsgs) {
        if (CollUtil.isEmpty(errorMsgs)) {
            return;
        }
        for (String errorMsg : errorMsgs) {
            addErrorInfo(fieldName, errorMsg);
        }
    }

    public void addVerifyResult(VerifyResult verifyResult) {
        if (Objects.isNull(verifyResult) || CollUtil.isEmpty(verifyResult.errorInfoMap)) {
            return;
        }
        for (Map.Entry<String, List<String>> entry : verifyResult.errorInfoMap.entrySet()) {
            addErrorInfo(entry.getKey(), entry.getValue());
        }
    }

    public void addVerifyResult(VerifyResult verifyResult, String parentFieldName, String parentFieldNameSuffix) {
        if (Objects.isNull(verifyResult) || CollUtil.isEmpty(verifyResult.errorInfoMap)) {
            return;
        }
        StrBuilder fieldName = StrUtil.strBuilder();
        for (Map.Entry<String, List<String>> entry : verifyResult.errorInfoMap.entrySet()) {
            fieldName.append(parentFieldName).append(parentFieldNameSuffix).append(entry.getKey());
            addErrorInfo(fieldName.toStringAndReset(), entry.getValue());
        }
    }

    public boolean hasError() {
        if (CollUtil.isEmpty(errorInfoMap)) {
            return false;
        }
        return errorInfoMap.values().stream().anyMatch(CollUtil::isNotEmpty);
    }

}
