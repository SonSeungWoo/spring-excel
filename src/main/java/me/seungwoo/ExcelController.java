package me.seungwoo;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;

import javax.servlet.http.HttpServletResponse;
import java.util.Arrays;
import java.util.List;

/**
 * Created by Leo.
 * User: ssw
 * Date: 2019-04-30
 * Time: 16:35
 */
@Controller
public class ExcelController {
    private void populateModel(Model model) {
        List<Person> rows = Arrays.asList(
                new Person("사람1", 10, "test@test.com"),
                new Person("사람2", 20, "test@test.com"),
                new Person("사람3", 30, "test@test.com")
        );
        model.addAttribute("rows", rows);

    }

    @GetMapping("/person.xls")
    public String getExcelByExt(Model model) {
        populateModel(model);
        return "personXls";
    }

    @GetMapping(path = "/person", params = "format=xls")
    public String getExcelByParam(Model model) {
        populateModel(model);
        return "personXls";
    }

    @GetMapping("/excel/download")
    public void excelDown(HttpServletResponse response) throws Exception {
        // DB조회
        List<Person> list = Arrays.asList(
                new Person("사람1", 10, "test@test.com"),
                new Person("사람2", 20, "test@test.com"),
                new Person("사람3", 30, "test@test.com")
        );
        //Workbook wb = new HSSFWorkbook(); //xls
        Workbook wb = new XSSFWorkbook();   //xlsx
        Sheet sheet = wb.createSheet("테스트");
        Row row;
        Cell cell;
        int rowNo = 0;

        // 테이블 헤더 스타일
        CellStyle headStyle = wb.createCellStyle();
        headStyle.setBorderTop(BorderStyle.THIN);
        headStyle.setBorderBottom(BorderStyle.THIN);
        headStyle.setBorderLeft(BorderStyle.THIN);
        headStyle.setBorderRight(BorderStyle.THIN);
        // 헤더 배경색 설정
        headStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.YELLOW.getIndex());
        headStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headStyle.setAlignment(HorizontalAlignment.CENTER);

        //테두리 지정
        CellStyle bodyStyle = wb.createCellStyle();
        bodyStyle.setBorderTop(BorderStyle.THIN);
        bodyStyle.setBorderBottom(BorderStyle.THIN);
        bodyStyle.setBorderLeft(BorderStyle.THIN);
        bodyStyle.setBorderRight(BorderStyle.THIN);

        // 헤더 추가 생성
        row = sheet.createRow(rowNo++);
        cell = row.createCell(0);
        cell.setCellStyle(headStyle);
        cell.setCellValue("이름");
        cell = row.createCell(1);
        cell.setCellStyle(headStyle);
        cell.setCellValue("나이");
        cell = row.createCell(2);
        cell.setCellStyle(headStyle);
        cell.setCellValue("이메일");

        // 데이터 부분 생성 추가로 생성
        for (Person person : list) {
            row = sheet.createRow(rowNo++);
            cell = row.createCell(0);
            cell.setCellStyle(bodyStyle);
            cell.setCellValue(person.getName());
            cell = row.createCell(1);
            cell.setCellStyle(bodyStyle);
            cell.setCellValue(person.getAge());
            cell = row.createCell(2);
            cell.setCellStyle(bodyStyle);
            cell.setCellValue(person.getEmail());
        }

        response.setContentType("ms-vnd/excel");
        response.setHeader("Content-Disposition", "attachment;filename=person.xlsx");
        wb.write(response.getOutputStream());
        wb.close();
    }
}
