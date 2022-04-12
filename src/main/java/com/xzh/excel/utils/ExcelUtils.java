package com.xzh.excel.utils;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.xzh.excel.model.ExcelObject;
import lombok.extern.slf4j.Slf4j;
import org.springframework.web.context.request.RequestAttributes;
import org.springframework.web.context.request.RequestContextHolder;
import org.springframework.web.context.request.ServletRequestAttributes;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * @author 向振华
 * @date 2021/04/22 10:47
 */
@Slf4j
public class ExcelUtils {

    /**
     * 将Excel文件解析成List
     *
     * @param file 文件
     * @param head 表头
     * @param <T>  泛型
     */
    public static <T> List<T> resolve(MultipartFile file, Class<T> head) {
        List<T> list = new ArrayList<>();
        AnalysisEventListener<T> analysisEventListener = new AnalysisEventListener<T>() {
            @Override
            public void invoke(T data, AnalysisContext context) {
                log.info("解析到一条数据:{}", data);
                list.add(data);
            }

            @Override
            public void doAfterAllAnalysed(AnalysisContext context) {
                log.info("所有数据解析完成！");
            }
        };
        try {
            EasyExcel.read(file.getInputStream(), head, analysisEventListener).sheet().doRead();
        } catch (Exception e) {
            log.error("upload-Exception：", e);
        }
        return list;
    }

    /**
     * Excel简单导出
     *
     * @param excelName 文件名
     * @param data      数据
     * @param head      表头
     * @param <T>       泛型
     */
    public static <T> void export(String excelName, List<T> data, Class<T> head) {
        try {
            ServletOutputStream outputStream = getOutputStream(excelName);
            EasyExcel.write(outputStream, head).sheet().doWrite(data);
        } catch (Exception e) {
            log.error("download-Exception：", e);
        }
    }

    /**
     * Excel多工作表导出
     *
     * @param excelName    文件名
     * @param excelObjects 文件对象
     */
    public static void exports(String excelName, List<ExcelObject> excelObjects) {
        ExcelWriter excelWriter = null;
        try {
            ServletOutputStream outputStream = getOutputStream(excelName);
            excelWriter = EasyExcel.write(outputStream).build();
            for (ExcelObject excelObject : excelObjects) {
                WriteSheet writeSheet = EasyExcel.writerSheet(excelObject.getSheet()).head(excelObject.getHead()).build();
                excelWriter.write(excelObject.getData(), writeSheet);
            }
        } catch (Exception e) {
            log.error("download-Exception：", e);
        } finally {
            if (excelWriter != null) {
                excelWriter.finish();
            }
        }
    }

    private static ServletOutputStream getOutputStream(String excelName) throws Exception {
        RequestAttributes requestAttributes = RequestContextHolder.getRequestAttributes();
        ServletRequestAttributes servletRequestAttributes = (ServletRequestAttributes) requestAttributes;
        assert servletRequestAttributes != null;
        HttpServletResponse response = servletRequestAttributes.getResponse();
        assert response != null;
        response.setContentType("application/vnd.ms-excel");
        response.setCharacterEncoding("UTF-8");
        response.setHeader("Access-Control-Expose-Headers", "Content-Disposition");
        response.setHeader("Content-Disposition", "attachment;filename=" + getFileName(excelName));
        return response.getOutputStream();
    }

    private static String getFileName(String excelName) {
        String fileName = new SimpleDateFormat("yyyyMMddHHmmss").format(new Date()) + ExcelTypeEnum.XLSX.getValue();
        try {
            return URLEncoder.encode(excelName, "UTF-8") + fileName;
        } catch (UnsupportedEncodingException e) {
            return fileName;
        }
    }
}
