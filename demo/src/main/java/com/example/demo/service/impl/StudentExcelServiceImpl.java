package com.example.demo.service.impl;

import com.example.demo.dto.StudentDTO;
import com.example.demo.service.StudentExcelService;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;


public class StudentExcelServiceImpl implements StudentExcelService {

    private XSSFWorkbook workbook;
    private XSSFSheet sheet;

    private List<StudentDTO> studentDTOs;


    public StudentExcelServiceImpl() {
        this.studentDTOs = getData();
        this.workbook = new XSSFWorkbook();
    }

    private List<StudentDTO> getData() {

        StudentDTO lahiru = new StudentDTO();
        lahiru.setId(1L);
        lahiru.setName("Lahiru");
        lahiru.setAddress("102 Bui Vien Street");
        lahiru.setCity("Sai Gon");
        lahiru.setPin("1234");

        StudentDTO envy = new StudentDTO();
        envy.setId(2L);
        envy.setName("Envy");
        envy.setAddress("99 Tran Hung Dao Street");
        envy.setCity("Ha Noi");
        envy.setPin("5678");

        List<StudentDTO> studentDTOs = Arrays.asList(lahiru, envy);

        return studentDTOs;
    }

    private void createCell(Row row, int columnCount, Object value, CellStyle style) {

        this.sheet.autoSizeColumn(columnCount);
        Cell cell = row.createCell(columnCount);

        if(value instanceof Long) {
            cell.setCellValue((Long) value);
        }else if(value instanceof Integer) {
            cell.setCellValue((Integer) value);
        }else if(value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
        }else {
            cell.setCellValue((String) value);
        }
        cell.setCellStyle(style);
    }

    private void writeHeaderLine() {

        this.sheet = this.workbook.createSheet("Student");

        Row row = this.sheet.createRow(0);

        CellStyle style = this.workbook.createCellStyle();
        XSSFFont font = this.workbook.createFont();
        font.setBold(true);
        font.setFontHeight(20);
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);

        createCell(row,0,"Student Information",style);
        this.sheet.addMergedRegion(new CellRangeAddress(0,0,0,4));
        //font.setFontHeightInPoints((short)(10));

        row = this.sheet.createRow(1);
        font.setBold(true);
        font.setFontHeight(16);
        style.setFont(font);

        createCell(row, 0, "Student Id", style);
        createCell(row, 1, "Student Name", style);
        createCell(row, 2, "Student Address", style);
        createCell(row, 3, "Student City", style);
        createCell(row, 4, "Student Pin", style);
    }

    private void writeDataLines() {

        int rowCount=2;

        CellStyle style = this.workbook.createCellStyle();
        XSSFFont font = this.workbook.createFont();
        font.setFontHeight(14);
        style.setFont(font);

        for(StudentDTO studentDTO : this.studentDTOs ) {

            Row row = this.sheet.createRow(rowCount++);
            int columnCount=0;

            createCell(row, columnCount++, studentDTO.getId(), style);
            createCell(row, columnCount++, studentDTO.getName(), style);
            createCell(row, columnCount++, studentDTO.getAddress(), style);
            createCell(row, columnCount++, studentDTO.getCity(), style);
            createCell(row, columnCount++, studentDTO.getPin(), style);
        }
    }

    @Override
    public void exportToExcel(HttpServletResponse response) throws IOException {

        writeHeaderLine();
        writeDataLines();

        ServletOutputStream outputStream=response.getOutputStream();
        this.workbook.write(outputStream);
        this.workbook.close();
        outputStream.close();
    }

}
