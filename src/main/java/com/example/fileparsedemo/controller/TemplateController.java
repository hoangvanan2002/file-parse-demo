package com.example.fileparsedemo.controller;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;

@Controller
public class TemplateController {

    @GetMapping("/file")
    public String indexTemplate(){
        return "index.html";
    }
}
