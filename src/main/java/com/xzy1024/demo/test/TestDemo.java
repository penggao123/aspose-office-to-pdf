package com.xzy1024.demo.test;

import com.xzy1024.demo.utils.PDFUtil;
import org.springframework.web.bind.annotation.*;

import javax.servlet.http.HttpServletRequest;

@RestController
@RequestMapping("/test")
public class TestDemo {


    @PostMapping("/imgToPdf")
    public String imgToPdf(@RequestParam("path") String path,@RequestParam("p") String p,  HttpServletRequest request){

//        return PDFUtil.convert2PDF(path, request);
        return null;
    }
    @PostMapping("/excToPdf")
    public boolean excToPdf(@RequestParam("path") String path,@RequestParam("p") String p,  HttpServletRequest request){

        return PDFUtil.excel2PDF(path, p);
    }
    @PostMapping("/docToPdf")
    public boolean docToPdf(@RequestParam("path") String path,@RequestParam("p") String p,  HttpServletRequest request){

        return PDFUtil.word2PDF(path, p);
    }

    @PostMapping("/pptToPdf")
    public boolean pptToPdf(@RequestParam("path") String path,@RequestParam("p") String p,  HttpServletRequest request) {

        return PDFUtil.ppt2PDF(path, p);
    }

    @PostMapping("/imageToPdf")
    public boolean imageToPdf(@RequestParam("path") String path,@RequestParam("p") String p,  HttpServletRequest request) {

        return PDFUtil.imgOfPdf(path, p);
    }
}
