package com.ncsi;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;

@Controller
public class PageController {

    @GetMapping("/public-report")
    public String publicReport() {
        return "forward:/public-report.html";
    }

    @GetMapping("/private-report")
    public String privateReport() {
        return "forward:/private-report.html";
    }
}


