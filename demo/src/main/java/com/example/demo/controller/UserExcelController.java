package com.example.demo.controller;

import com.example.demo.service.StudentExcelService;
import com.example.demo.service.impl.StudentExcelServiceImpl;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

@RestController
@RequestMapping("api/v1/")
public class UserExcelController {


    @GetMapping("student/excel")
    public void exportToExcel(HttpServletResponse response) throws IOException {

        response.setContentType("application/octet-stream");
        DateFormat dateFormatter = new SimpleDateFormat("yyyy-MM-dd_HH-mm-ss");
        String currentDateTime = dateFormatter.format(new Date());

        String headerKey = "Content-Disposition";
        String headerValue = "attachment; filename=Student_Info_" + currentDateTime + ".xlsx";
        response.setHeader(headerKey, headerValue);

        StudentExcelService studentExcelService = new StudentExcelServiceImpl();
        studentExcelService.exportToExcel(response);
    }

}
