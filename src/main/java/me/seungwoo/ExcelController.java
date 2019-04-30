package me.seungwoo;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
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

    @GetMapping("/excelDown")
    public void excelDown(HttpServletResponse response) throws Exception {
        // 게시판 목록조회
        List<Person> list = Arrays.asList(
                new Person("사람1", 10, "test@test.com"),
                new Person("사람2", 20, "test@test.com"),
                new Person("사람3", 30, "test@test.com")
        );

        // 워크북 생성
        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("게시판");
        Row row;
        Cell cell;
        int rowNo = 0;

        // 테이블 헤더용 스타일
        CellStyle headStyle = wb.createCellStyle();

        // 가는 경계선을 가집니다.
        headStyle.setBorderTop(BorderStyle.THIN);
        headStyle.setBorderBottom(BorderStyle.THIN);
        headStyle.setBorderLeft(BorderStyle.THIN);
        headStyle.setBorderRight(BorderStyle.THIN);

        // 배경색은 노란색입니다.
        headStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.YELLOW.getIndex());
        headStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // 데이터는 가운데 정렬합니다.
        headStyle.setAlignment(HorizontalAlignment.CENTER);

        // 데이터용 경계 스타일 테두리만 지정
        CellStyle bodyStyle = wb.createCellStyle();
        bodyStyle.setBorderTop(BorderStyle.THIN);
        bodyStyle.setBorderBottom(BorderStyle.THIN);
        bodyStyle.setBorderLeft(BorderStyle.THIN);
        bodyStyle.setBorderRight(BorderStyle.THIN);

        // 헤더 생성
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

        // 데이터 부분 생성
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

        // 컨텐츠 타입과 파일명 지정
        response.setContentType("ms-vnd/excel");
        response.setHeader("Content-Disposition", "attachment;filename=person.xls");

        // 엑셀 출력
        wb.write(response.getOutputStream());
        wb.close();
    }
}
