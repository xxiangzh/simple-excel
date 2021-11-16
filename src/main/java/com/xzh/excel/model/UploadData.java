package com.xzh.excel.model;

import com.alibaba.excel.annotation.ExcelProperty;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * @author 向振华
 * @date 2021/04/22 10:24
 */
@Data
@NoArgsConstructor
@AllArgsConstructor
public class UploadData {

    @ExcelProperty(value = "编码")
    private Long id;

    @ExcelProperty(value = "名称")
    private String name;
}
