package me.seungwoo;

import org.apache.poi.ss.usermodel.*;
import org.springframework.context.annotation.Configuration;
import org.springframework.web.servlet.view.document.AbstractXlsView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.List;
import java.util.Map;

/**
 * Created by Leo.
 * User: ssw
 * Date: 2019-04-30
 * Time: 16:41
 */

@Configuration("personXls")
public class ExcelView extends AbstractXlsView {
    @Override
    protected void buildExcelDocument(Map<String, Object> map, Workbook workbook,
                                      HttpServletRequest httpServletRequest,
                                      HttpServletResponse httpServletResponse) throws Exception {
        httpServletResponse.setHeader("Content-Disposition", "attachment; filename=\"person.xls\"");
        List<Person> persons = (List<Person>) map.get("rows");

        CellStyle numberCellStyle = workbook.createCellStyle();
        DataFormat numberDataFormat = workbook.createDataFormat();
        numberCellStyle.setDataFormat(numberDataFormat.getFormat("#,##0"));

        Sheet sheet = workbook.createSheet("model_person");
        for (int i = 0; i < persons.size(); i++) {
            Person person = persons.get(i);
            Row row = sheet.createRow(i);

            Cell cell0 = row.createCell(0);
            cell0.setCellValue(person.getName());

            Cell cell1 = row.createCell(1);
            cell1.setCellType(CellType.NUMERIC);
            cell1.setCellValue(person.getAge());
            cell1.setCellStyle(numberCellStyle);

            Cell cell2 = row.createCell(2);
            cell2.setCellType(CellType.NUMERIC);
            cell2.setCellValue(person.getEmail());
            cell2.setCellStyle(numberCellStyle);
        }
    }
}
