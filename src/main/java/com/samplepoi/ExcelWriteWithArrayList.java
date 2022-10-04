package com.samplepoi;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import static java.lang.System.*;

public class ExcelWriteWithArrayList {

    public static void main(String[] args) throws IOException {

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Emp info");

        ArrayList<Object[]> empData = new ArrayList<Object[]>();
        empData.add(new Object[]{"EmpId", "Name", "Job"});
        empData.add(new Object[]{101, "Nivesh", "Engineer"});
        empData.add(new Object[]{102, "Ram", "Doctor"});
        empData.add(new Object[]{103, "Gopal", "Teacher"});
        empData.add(new Object[]{104, "Ramesh", "Analyst"});
        empData.add(new Object[]{105, "Venu", "Accountant"});

        // Using for each loop

        int rowCount = 0;

        for (Object emp[] : empData) {
            XSSFRow row = sheet.createRow(rowCount++);
            int colCount = 0;
            for (Object value : emp) {
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
        FileOutputStream fileOutputStream = new FileOutputStream(filePath);
        workbook.write(fileOutputStream);

        fileOutputStream.close();
        out.println("employee.xlsx file written successfully");
    }
}
