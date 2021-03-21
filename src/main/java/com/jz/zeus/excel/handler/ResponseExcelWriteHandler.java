package com.jz.zeus.excel.handler;

import com.jz.zeus.excel.ExcelInfo;
import lombok.SneakyThrows;

import javax.servlet.http.HttpServletResponse;
import java.io.OutputStream;
import java.net.URLEncoder;

/**
 * @author:JZ
 * @date:2021/3/21
 */
public class ResponseExcelWriteHandler implements ExcelWriteHandler<HttpServletResponse> {

    private OutputStreamExcelWriteHandler excelWriteHandler = new OutputStreamExcelWriteHandler();

    @Override
    public boolean check(ExcelInfo excelInfo) {
        return excelWriteHandler.check(excelInfo);
    }

    @Override
    @SneakyThrows
    public void execute(ExcelInfo excelInfo, HttpServletResponse response) {
        check(excelInfo);
        response.setContentType("application/vnd.ms-excel");
        response.setCharacterEncoding("utf-8");
        String fileName = URLEncoder.encode(excelInfo.getName(), "UTF-8").replaceAll("\\+", "%20");
        response.setHeader("Content-disposition", "attachment;filename*=utf-8''" + fileName + excelInfo.getSuffix().name());
        write(excelInfo, response);
    }

    @Override
    @SneakyThrows
    public void write(ExcelInfo excelInfo, HttpServletResponse response) {
        excelWriteHandler.write(excelInfo, response.getOutputStream());
    }
}
