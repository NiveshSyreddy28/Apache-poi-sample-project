package com.samplepoi;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

public class ExcelRead {
    private static XSSFWorkbook workbook;
    private static XSSFSheet sheet;
    private static FileInputStream fileInputStream;


    public static void main(String[] args) throws IOException {
        String filePath = "/home/niveS/Desktop/SampleApachePoIProject/datafiles/testdata.xlsx";
        fileInputStream = new FileInputStream(filePath);

        workbook = new XSSFWorkbook(fileInputStream);
        sheet = workbook.getSheetAt(0);

        // Using For Loop

        /*int noOfRows = sheet.getLastRowNum();
        System.out.println(noOfRows);

        int noOfCols = sheet.getRow(1).getLastCellNum();
        System.out.println(noOfCols);

        for(int r =0; r <= noOfRows; r++){

            XSSFRow row = sheet.getRow(r);

            for (int c = 0; c < noOfCols; c++){

                XSSFCell cell = row.getCell(c);

                switch (cell.getCellType()){

                    case STRING:
                        System.out.print(cell.getStringCellValue());
                        break;
                    case NUMERIC:
                        System.out.print(cell.getNumericCellValue());
                        break;
                    case BOOLEAN:
                        System.out.print(cell.getBooleanCellValue());
                        break;
                }
                System.out.print(" | ");
            }
            System.out.println();
        }*/

        //Using Iterator

        Iterator iterator = sheet.iterator();

        while (iterator.hasNext()){
            XSSFRow row = (XSSFRow) iterator.next();

           Iterator cellIterator = row.cellIterator();

           while (cellIterator.hasNext()){
               XSSFCell cell = (XSSFCell) cellIterator.next();

               switch (cell.getCellType()){

                   case STRING:
                       System.out.print(cell.getStringCellValue());
                       break;
                   case NUMERIC:
                       System.out.print(cell.getNumericCellValue());
                       break;
                   case BOOLEAN:
                       System.out.print(cell.getBooleanCellValue());
                       break;
               }
               System.out.print(" | ");
           }
            System.out.println();
        }
    }
}
