// Corrected: FrontendController.java
package com.example.certgenerator.demo.controller;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

@Controller
public class FrontendController {

    @RequestMapping(value = {"/{path:[^\\.]*}", "/"})
    public String redirect() {
        return "forward:/index.html";
    }
}

