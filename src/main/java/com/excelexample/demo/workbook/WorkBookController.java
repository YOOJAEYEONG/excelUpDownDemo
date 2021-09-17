package com.excelexample.demo.workbook;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;


@Controller
public class WorkBookController {

    @GetMapping("/excel/down")
    public void getExcel(HttpServletResponse hsr) throws IOException {
        System.out.println("getExcel");
//        xls 확장자를 가진 엑셀 파일 -> HSSFWorkbook
//        xlsx 확장자를 가진 엑셀 파일 -> XSSFWorkbook
        XSSFWorkbook wb = new XSSFWorkbook();
        Sheet s = wb.createSheet("시트이름");
        Row r;
        Cell c;
        int rowNum = 0;


        //header
        r = s.createRow(rowNum++);
        c = r.createCell(0);
        c.setCellValue("번호");

        c = r.createCell(1);
        c.setCellValue("이름");
        c = r.createCell(2);
        c.setCellValue("제목");
        CellStyle style1 = wb.createCellStyle();
//        style1.setBottomBorderColor();

        //body
        for (int i=0; i<10; i++){
            r = s.createRow(rowNum++);
            c = r.createCell(0);
            c.setCellValue(i);
            c = r.createCell(1);
            c.setCellValue(i+"name");
            c = r.createCell(2);
            c.setCellValue(i+"title");
        }

        //컨텐츠 타입과 파일명지정
        hsr.setContentType("ms-vnd/excel");
        hsr.setHeader("Content-Disposition", "attachment;filename=example.xlsx");

        //Excel File Output
        wb.write(hsr.getOutputStream());
        wb.close();
    }
}
