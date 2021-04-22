package com.xzh.excel.controller;

import com.xzh.excel.model.UploadData;
import com.xzh.excel.utils.ExcelUtils;
import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;
import io.swagger.annotations.ApiParam;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.util.ArrayList;
import java.util.List;

/**
 * @author 向振华
 * @date 2021/04/22 09:04
 */
@Api(tags = "表格")
@RestController
@RequestMapping("/excel")
public class ExcelController {

    @ApiOperation("表格上传")
    @PostMapping("/upload")
    public String upload(@ApiParam("file") MultipartFile file) {
        List<UploadData> upload = ExcelUtils.upload(file, UploadData.class);
        return "success";
    }

    @ApiOperation("表格下载")
    @PostMapping("/download")
    public String download(HttpServletResponse response) {
        List<UploadData> data = new ArrayList<>();
        data.add(new UploadData(1L, "哈哈"));
        data.add(new UploadData(3L, "dfshg"));
        data.add(new UploadData(5L, "2333"));
        ExcelUtils.download(response, "我是文件名", "我是工作表名称", data, UploadData.class);
        return "success";
    }
}