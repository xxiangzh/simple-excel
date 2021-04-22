package com.xzh.excel.controller;

import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

/**
 * @author 向振华
 * @date 2021/04/22 09:04
 */
@Api(tags = "主页")
@RestController
@RequestMapping
public class HomeController {

    @ApiOperation("主页")
    @GetMapping
    public String home() {
        return "success";
    }
}