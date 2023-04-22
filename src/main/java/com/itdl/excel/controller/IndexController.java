package com.itdl.excel.controller;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;

/**
 * @Description
 * @Author IT动力
 * @Date 2023/04/20 11:39
 */
@Controller
public class IndexController {

    @GetMapping("")
    public String index(){
        return "download.html";
    }

}
