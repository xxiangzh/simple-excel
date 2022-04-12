package com.xzh.excel.model;

import lombok.Data;

import java.util.List;

/**
 * @author 向振华
 * @date 2021/08/16 09:17
 */
@Data
public class ExcelObject {

    /**
     * 工作表名
     */
    private String sheet;

    /**
     * 表数据
     */
    private List<?> data;

    /**
     * 表头
     */
    private Class<?> head;
}