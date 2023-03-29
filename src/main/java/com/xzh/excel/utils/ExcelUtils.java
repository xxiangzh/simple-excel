package com.xzh.excel.utils;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.exception.ExcelAnalysisException;
import com.alibaba.excel.exception.ExcelDataConvertException;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.fastjson.JSONObject;
import com.xzh.excel.model.ExcelObject;
import lombok.extern.slf4j.Slf4j;
import org.springframework.web.context.request.RequestAttributes;
import org.springframework.web.context.request.RequestContextHolder;
import org.springframework.web.context.request.ServletRequestAttributes;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
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
     * @param pathName D:\demo.xlsx
     * @param head     表头
     * @param <T>      泛型
     * @return
     */
    public static <T> List<T> read(String pathName, Class<T> head) {
        List<T> list = new ArrayList<>();
        try {
            EasyExcel.read(pathName, head, new AnalysisEventListener<T>() {
                @Override
                public void invoke(T data, AnalysisContext context) {
                    log.info("解析到一条数据 " + JSONObject.toJSONString(data));
                    list.add(data);
                }

                @Override
                public void doAfterAllAnalysed(AnalysisContext context) {
                    log.info("所有数据解析完成 " + list.size());
                }
            }).sheet().doRead();
        } catch (Exception e) {
            log.error("readExcelException ", e);

            ExcelDataConvertException ex = null;
            if (e instanceof ExcelDataConvertException) {
                ex = (ExcelDataConvertException) e;
            } else if (e instanceof ExcelAnalysisException) {
                ExcelAnalysisException excelAnalysisException = (ExcelAnalysisException) e;
                Throwable cause = excelAnalysisException.getCause();
                if (cause instanceof ExcelDataConvertException) {
                    ex = (ExcelDataConvertException) cause;
                }
            }

            if (ex != null) {
                String errorMsg = String.format("解析文件异常：第%d行 %s 格式错误", ex.getRowIndex(), ex.getCellData());
                throw new RuntimeException(errorMsg);
            } else {
                throw e;
            }
        }
        return list;
    }

    /**
     * 将Excel文件解析成List
     *
     * @param file 文件
     * @param head 表头
     * @param <T>  泛型
     */
    public static <T> List<T> read(MultipartFile file, Class<T> head) {
        List<T> list = new ArrayList<>();
        try {
            EasyExcel.read(file.getInputStream(), head, new AnalysisEventListener<T>() {
                @Override
                public void invoke(T data, AnalysisContext context) {
                    log.info("解析到一条数据 " + JSONObject.toJSONString(data));
                    list.add(data);
                }

                @Override
                public void doAfterAllAnalysed(AnalysisContext context) {
                    log.info("所有数据解析完成 " + list.size());
                }
            }).sheet().doRead();
        } catch (Exception e) {
            log.error("readExcelException ", e);

            ExcelDataConvertException ex = null;
            if (e instanceof ExcelDataConvertException) {
                ex = (ExcelDataConvertException) e;
            } else if (e instanceof ExcelAnalysisException) {
                ExcelAnalysisException excelAnalysisException = (ExcelAnalysisException) e;
                Throwable cause = excelAnalysisException.getCause();
                if (cause instanceof ExcelDataConvertException) {
                    ex = (ExcelDataConvertException) cause;
                }
            }

            if (ex != null) {
                String errorMsg = String.format("解析文件异常：第%d行 %s 格式错误", ex.getRowIndex(), ex.getCellData());
                throw new RuntimeException(errorMsg);
            } else {
                throw new RuntimeException("解析文件异常：" + e.getMessage());
            }
        }
        return list;
    }

    /**
     * Excel导出到本地
     *
     * @param targetFolderDirectory 目标文件夹目录 D:\xxx
     * @param fileName              文件名
     * @param data                  数据
     * @param head                  表头
     * @param <T>                   泛型
     */
    public static <T> void write(String targetFolderDirectory, String fileName, List<T> data, Class<T> head) {
        fileName = rectifyFileName(fileName);
        try {
            EasyExcel.write(targetFolderDirectory + fileName, head).sheet().doWrite(data);
        } catch (Exception e) {
            log.error("writeExcelException ", e);
        }
    }

    /**
     * Excel简单导出
     *
     * @param fileName 文件名
     * @param data     数据
     * @param head     表头
     * @param <T>      泛型
     */
    public static <T> void write(String fileName, List<T> data, Class<T> head) {
        fileName = rectifyFileName(fileName);
        try {
            ServletOutputStream outputStream = getOutputStream(fileName);
            EasyExcel.write(outputStream, head).sheet().doWrite(data);
        } catch (Exception e) {
            log.error("writeExcelException ", e);
        }
    }

    /**
     * Excel多工作表导出
     *
     * @param fileName     文件名
     * @param excelObjects 文件对象
     */
    public static void write(String fileName, List<ExcelObject> excelObjects) {
        fileName = rectifyFileName(fileName);
        ExcelWriter excelWriter = null;
        try {
            ServletOutputStream outputStream = getOutputStream(fileName);
            excelWriter = EasyExcel.write(outputStream).build();
            for (ExcelObject excelObject : excelObjects) {
                WriteSheet writeSheet = EasyExcel.writerSheet(excelObject.getSheet()).head(excelObject.getHead()).build();
                excelWriter.write(excelObject.getData(), writeSheet);
            }
        } catch (Exception e) {
            log.error("writeExcelException ", e);
        } finally {
            if (excelWriter != null) {
                excelWriter.finish();
            }
        }
    }

    /**
     * 矫正文件名
     *
     * @param fileName
     * @return
     */
    private static String rectifyFileName(String fileName) {
        String time = "_" + new SimpleDateFormat("yyyyMMddHHmmss").format(new Date());
        if (fileName == null || fileName.isEmpty()) {
            fileName = "";
        }
        if (fileName.endsWith(ExcelTypeEnum.XLS.getValue())) {
            fileName = fileName.replace(ExcelTypeEnum.XLS.getValue(), time + ExcelTypeEnum.XLS.getValue());
        } else if (fileName.endsWith(ExcelTypeEnum.XLSX.getValue())) {
            fileName = fileName.replace(ExcelTypeEnum.XLSX.getValue(), time + ExcelTypeEnum.XLSX.getValue());
        } else {
            fileName = fileName + time + ExcelTypeEnum.XLSX.getValue();
        }
        return fileName;
    }

    /**
     * 获取输出流
     *
     * @param fileName
     * @return
     * @throws Exception
     */
    private static ServletOutputStream getOutputStream(String fileName) throws Exception {
        RequestAttributes requestAttributes = RequestContextHolder.getRequestAttributes();
        ServletRequestAttributes servletRequestAttributes = (ServletRequestAttributes) requestAttributes;
        assert servletRequestAttributes != null;
        HttpServletResponse response = servletRequestAttributes.getResponse();
        assert response != null;
        response.setContentType("application/vnd.ms-excel");
        response.setCharacterEncoding("UTF-8");
        response.setHeader("Access-Control-Expose-Headers", "Content-Disposition");
        response.setHeader("Content-Disposition", "attachment;filename=" + URLEncoder.encode(fileName, "UTF-8"));
        return response.getOutputStream();
    }
}
