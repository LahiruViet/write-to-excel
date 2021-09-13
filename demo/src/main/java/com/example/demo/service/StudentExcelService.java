package com.example.demo.service;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;

public interface StudentExcelService {

    void exportToExcel(HttpServletResponse response) throws IOException;
}
