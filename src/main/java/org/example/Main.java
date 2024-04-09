package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;

// Press Shift twice to open the Search Everywhere dialog and type `show whitespaces`,
// then press Enter. You can now see whitespace characters in your code.
public class Main {
    public static void main(String[] args)
    {

        // Try block to check for exceptions 
        try {

            // Reading file from local directory 
            FileInputStream file = new FileInputStream(new File("C:\\Users\\senon\\Desktop\\Book1.xlsx"));

            // Create Workbook instance holding reference to 
            // .xlsx file 
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            // Get first/desired sheet from the workbook 
            XSSFSheet sheet = workbook.getSheetAt(0);

            // Iterate through each rows one by one 
            Iterator<Row> rowIterator = sheet.iterator();

            // Till there is an element condition holds true 
            while (rowIterator.hasNext()) {

                Row row = rowIterator.next();

                // For each row, iterate through all the 
                // columns 
                Iterator<Cell> cellIterator
                        = row.cellIterator();

                while (cellIterator.hasNext()) {

                    Cell cell = cellIterator.next();

                    // Checking the cell type and format 
                    // accordingly 
                    switch (cell.getCellType()) {

                        // Case 1 
                        case NUMERIC:
                            System.out.print(
                                    cell.getNumericCellValue()
                                            + "t");
                            break;

                        // Case 2 
                        case STRING:
                            System.out.print(
                                    cell.getStringCellValue()
                                            + "t");
                            break;
                    }
                }

                System.out.println("");
            }

            // Closing file output streams 
            file.close();
        }

        // Catch block to handle exceptions 
        catch (Exception e) {

            // Display the exception along with line number 
            // using printStackTrace() method 
            e.printStackTrace();
        }
    }
}