package com.xzh.excel.utils;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.fastjson.JSON;
import lombok.extern.slf4j.Slf4j;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.List;

/**
 * @author 向振华
 * @date 2021/04/22 10:47
 */
@Slf4j
public class ExcelUtils {

    /**
     * 上传Excel文件，并解析成List
     *
     * @param file  文件
     * @param clazz 类
     * @param <T>   泛型
     * @return
     */
    public static <T> List<T> upload(MultipartFile file, Class<T> clazz) {
        List<T> list = new ArrayList<>();
        AnalysisEventListener<T> analysisEventListener = new AnalysisEventListener<T>() {
            @Override
            public void invoke(T data, AnalysisContext context) {
                log.info("解析到一条数据:{}", JSON.toJSONString(data));
                list.add(data);
            }

            @Override
            public void doAfterAllAnalysed(AnalysisContext context) {
                log.info("所有数据解析完成！");
            }
        };
        try {
            EasyExcel.read(file.getInputStream(), clazz, analysisEventListener).sheet().doRead();
        } catch (IOException e) {
            log.error("解析异常：", e);
        }
        return list;
    }

    /**
     * Excel文件下载
     *
     * @param response  响应流
     * @param excelName 文件名
     * @param sheetName 工作表名称
     * @param data      数据
     * @param clazz     类
     * @param <T>       泛型
     */
    public static <T> void download(HttpServletResponse response, String excelName, String sheetName, List<T> data, Class<T> clazz) {
        response.setContentType("application/vnd.ms-excel");
        response.setCharacterEncoding("UTF-8");
        try {
            String fileName = URLEncoder.encode(excelName, "UTF-8");
            response.setHeader("Content-disposition", "attachment;filename=" + fileName + ".xlsx");
            EasyExcel.write(response.getOutputStream(), clazz).sheet(sheetName).doWrite(data);
        } catch (IOException e) {
            log.error("下载异常：", e);
        }
    }
}
