package com.example.cache.demo.controller;

import com.example.cache.demo.service.GenerateExcel;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;

import static java.util.Objects.nonNull;
import static org.apache.poi.util.IOUtils.copy;
import static org.springframework.http.HttpHeaders.CONTENT_DISPOSITION;
import static org.springframework.http.MediaType.APPLICATION_JSON_VALUE;

@RestController
public class ExcelController {

    @Autowired
    private GenerateExcel generateExcel;

    @GetMapping
    public void getControlledReleaseRequests(HttpServletResponse response) throws Exception {
        ByteArrayOutputStream byteArrayOutputStream = generateExcel.process(generateExcel.mock());

        if (nonNull(byteArrayOutputStream)) {
            InputStream isOs = new ByteArrayInputStream(byteArrayOutputStream.toByteArray());
            copy(isOs, response.getOutputStream());
            response.setHeader(CONTENT_DISPOSITION, "attachment; filename=qrcode.xlsx");
            response.setContentType("application/vnd.ms-excel");
        } else {
            response.setContentLength(0);
            response.setContentType(APPLICATION_JSON_VALUE);
        }
    }

    @GetMapping("/import")
    public void teste2(HttpServletResponse response) throws Exception {
        ByteArrayOutputStream byteArrayOutputStream = generateExcel.process1();

        if (nonNull(byteArrayOutputStream)) {
            InputStream isOs = new ByteArrayInputStream(byteArrayOutputStream.toByteArray());
            copy(isOs, response.getOutputStream());
            response.setHeader(CONTENT_DISPOSITION, "attachment; filename=qrcode_new.xlsx");
            response.setContentType("application/vnd.ms-excel");
        } else {
            response.setContentLength(0);
            response.setContentType(APPLICATION_JSON_VALUE);
        }
    }
}
