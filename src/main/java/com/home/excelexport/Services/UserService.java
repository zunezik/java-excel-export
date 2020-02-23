package com.home.excelexport.Services;

import com.home.excelexport.Models.User;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.util.Arrays;
import java.util.List;

@Service
public class UserService {
    private List<User> users = Arrays.asList(
            new User("Anton", 21),
            new User("Ann", 20));

    public List<User> getUsers() {
        return this.users;
    }

    public ByteArrayInputStream exportExcel() throws IOException {
        return createExcel();
    }

    public ByteArrayInputStream createExcel() throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Users");

        writeHeaderLine(sheet);
        writeDataLines(this.users, workbook, sheet);

        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        workbook.write(outputStream);
        workbook.close();
        return new ByteArrayInputStream(outputStream.toByteArray());
    }

    private void writeHeaderLine(XSSFSheet sheet) {
        Row headerRow = sheet.createRow(0);

        Cell headerCell = headerRow.createCell(0);
        headerCell.setCellValue("Name");

        headerCell = headerRow.createCell(1);
        headerCell.setCellValue("Age");
    }

    private void writeDataLines(List<User> users, XSSFWorkbook workbook,
                                XSSFSheet sheet) {
        int rowNumber = 1;

        for (User user : users) {
            Row row = sheet.createRow(rowNumber++);
            int columnNumber = 0;

            Cell cell = row.createCell(columnNumber++);
            cell.setCellValue(user.getName());

            cell = row.createCell(columnNumber++);
            cell.setCellValue(user.getAge());

            cell = row.createCell(columnNumber);

            CellStyle cellStyle = workbook.createCellStyle();
            CreationHelper creationHelper = workbook.getCreationHelper();
            cellStyle.setDataFormat(creationHelper.createDataFormat().getFormat("yyyy-MM-dd HH:mm:ss"));
            cell.setCellStyle(cellStyle);
            sheet.autoSizeColumn(columnNumber++);

            cell.setCellValue(LocalDateTime.now());
        }
    }
}
