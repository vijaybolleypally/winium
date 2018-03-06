import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

public class LifeofPI {
    public static void main(String[] args) throws IOException, InvalidFormatException {

        String SAMPLE_XLSX_FILE_PATH = "C:\\Users\\E002961\\Desktop\\exs\\Actual.xlsx";
        String SAMPLE_XLSX_FILE_PATH1 = "C:\\Users\\E002961\\Desktop\\exs\\Expected.xlsx";

        // Creating a Workbook from an Excel file (.xls or .xlsx)
        Workbook workbook = WorkbookFactory.create(new File(SAMPLE_XLSX_FILE_PATH));
        Workbook workbook1 = WorkbookFactory.create(new File(SAMPLE_XLSX_FILE_PATH1));
        Workbook workbook2 = new XSSFWorkbook();
        Sheet sheetN = workbook2.createSheet("RES");

//        // Create a Font for styling header cells
//        Font headerFont = workbook2.createFont();
//        headerFont.setBold(true);
//        headerFont.setFontHeightInPoints((short) 14);
//        headerFont.setColor(IndexedColors.RED.getIndex());

        // Create a CellStyle with the font
//        CellStyle headerCellStyle = workbook2.createCellStyle();
//        headerCellStyle.setFont(headerFont);


        // Retrieving the number of sheets in the Workbook
        System.out.println("Workbook has " + workbook.getNumberOfSheets() + " Sheets : ");
        System.out.println("Workbook has " + workbook1.getNumberOfSheets() + " Sheets : ");
        /*
           =============================================================
           Iterating over all the sheets in the workbook (Multiple ways)
           =============================================================
        */

        // 1. You can obtain a sheetIterator and iterate over it
        Iterator<Sheet> sheetIterator = workbook.sheetIterator();
        System.out.println("Retrieving Sheets using Iterator");
        while (sheetIterator.hasNext()) {
            Sheet sheet = sheetIterator.next();
            System.out.println("=> " + sheet.getSheetName());
        }

        // 3. Or you can use a Java 8 forEach with lambda
        System.out.println("Retrieving Sheets using Java 8 forEach with lambda");
        workbook1.forEach(sheet -> {
            System.out.println("=> " + sheet.getSheetName());
        });

        /*
           ==================================================================
           Iterating over all the rows and columns in a Sheet (Multiple ways)
           ==================================================================
        */

        // Getting the Sheet at index zero
        Sheet sheet = workbook.getSheetAt(0);
        Sheet sheet1 = workbook1.getSheetAt(0);

        // Create a DataFormatter to format and get each cell's value as String
        DataFormatter dataFormatter = new DataFormatter();

        // 1. You can obtain a rowIterator and columnIterator and iterate over them
        System.out.println("\n\nIterating over Rows and Columns using Iterator\n");
        Iterator<Row> rowIterator = sheet.rowIterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();

            // Now let's iterate over the columns of the current row
            Iterator<Cell> cellIterator = row.cellIterator();

            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                String cellValue = dataFormatter.formatCellValue(cell);
                System.out.print(cellValue + "\t");
            }
            System.out.println();
        }

//         2. Or you can use a for-each loop to iterate over the rows and columns
        System.out.println("\n\n Write over Rows and Columns using for-each loop\n");
        for (int i = sheet.getFirstRowNum(); i <= sheet.getLastRowNum(); i++) {
            Row headerRow = sheetN.createRow(i);
            Row eachRow = sheet.getRow(i);
            for (int j = eachRow.getFirstCellNum(); j <= eachRow.getLastCellNum(); j++) {
                Cell eachCell = headerRow.createCell(j);
                String cellValueN = dataFormatter.formatCellValue(eachRow.getCell(j));
                eachCell.setCellValue(cellValueN);
                eachCell.setCellStyle(eachCell.getCellStyle());
            }
        }

        // 3. Or you can use Java 8 forEach loop with lambda
        System.out.println("\n\nIterating over Rows and Columns using Java 8 forEach with lambda\n");
        sheet1.forEach(row -> {
            row.forEach(cell -> {
                String cellValue = dataFormatter.formatCellValue(cell);
                System.out.print(cellValue + "\t");
            });
            System.out.println();
        });

        // Closing the workbook
        workbook.close();
        workbook1.close();
        FileOutputStream fileOut = new FileOutputStream("C:\\Users\\E002961\\Desktop\\exs\\Result.xlsx");
        workbook2.write(fileOut);
        fileOut.close();
    }
}
