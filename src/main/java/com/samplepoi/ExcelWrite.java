package com.samplepoi;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

import static java.lang.System.*;

public class ExcelWrite {
    public static void main(String[] args) throws IOException {

        FileOutputStream fileOutputStream;
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            XSSFSheet sheet = workbook.createSheet("Emp info");

            Object[][] empData = {
                    {"EmpId", "Name", "Job"},
                    {101, "Nivesh", "Engineer"},
                    {102, "Ram", "Doctor"},
                    {103, "Gopal", "Teacher"},
                    {104, "Ramesh", "Analyst"},
                    {105, "Venu", "Accountant"},
            };

            // Using For loop

            /* int rows = empData.length;
            int cols = empData[0].length;

            out.println(rows);
            out.println(cols);

            for (int r = 0; r < rows; r++) {
                XSSFRow row = sheet.createRow(r);

                for (int c = 0; c < cols; c++) {
                    XSSFCell cell = row.createCell(c);
                    Object value = empData[r][c];

                    if (value instanceof String) {
                        cell.setCellValue((String) value);
                    } else if (value instanceof Integer) {
                        cell.setCellValue((Integer) value);
                    } else if (value instanceof Boolean) {
                        cell.setCellValue((Boolean) value);
                    }
                }
            }*/

            // Using for each loop

            int rowCount = 0;

            for (Object emp[]:empData) {
                XSSFRow row = sheet.createRow(rowCount++);
                int colCount = 0;
                for (Object value: emp){
                    XSSFCell cell = row.createCell(colCount++);

                    if (value instanceof String) {
                        cell.setCellValue((String) value);
                    } else if (value instanceof Integer) {
                        cell.setCellValue((Integer) value);
                    } else if (value instanceof Boolean) {
                        cell.setCellValue((Boolean) value);
                    }
                }
            }
            String filePath = "/home/niveS/Desktop/SampleApachePoIProject/datafiles/employee.xlsx";
            fileOutputStream = new FileOutputStream(filePath);
            workbook.write(fileOutputStream);
        }

        fileOutputStream.close();
        out.println("employee.xlsx file written successfully");
    }
}
