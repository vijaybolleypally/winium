package util;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xpath.SourceTree;
import org.testng.Assert;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

public class SheetUtill {

    public Workbook getWorkBook(String excelFilePath) throws IOException, InvalidFormatException {
        return WorkbookFactory.create(new File(excelFilePath));
    }

    public void printSheetNameInWorkbook(Workbook givenWorkBook) {
//        System.out.println("Retrieving Sheets in "+givenWorkBook.;
    }

    public void compareExcel(String expectedFileWithPath, String actualFileWithPath, String resultFileWithPath) throws IOException, InvalidFormatException {
        //Get work book objects
        Workbook expectedWorkbook = getWorkBook(expectedFileWithPath);
        Workbook actualWorkbook = getWorkBook(actualFileWithPath);
        Workbook resultWorkbook = new XSSFWorkbook();

        //Print Sheets in Each Workbook
        System.out.println("\n\n Retrieving Sheets in:" + expectedFileWithPath);
        Iterator<Sheet> expectedSheetIterator = expectedWorkbook.sheetIterator();
        while (expectedSheetIterator.hasNext()) {
            Sheet sheet = expectedSheetIterator.next();
            System.out.println("=> " + sheet.getSheetName());
        }

        System.out.println(" \n\n Retrieving Sheets in:" + actualFileWithPath);
        Iterator<Sheet> actualSheetIterator = actualWorkbook.sheetIterator();
        while (actualSheetIterator.hasNext()) {
            Sheet sheet = actualSheetIterator.next();
            System.out.println("=> " + sheet.getSheetName());
        }

        //Compare Sheets Size in each excel and fail if not match
        Assert.assertEquals(expectedWorkbook.getNumberOfSheets(), actualWorkbook.getNumberOfSheets());

        int totalCellsCompared = 0, totalCellsMatched = 0, totalCellNotMatched = 0;

        //Go Through each sheet
        System.out.println("\n\nIterating over each sheet\n");
        for (int s = 0; s < expectedWorkbook.getNumberOfSheets(); s++) {
            System.out.println("\n\n Comparing Sheet at index : " + s);
            Sheet eachExpectedSheet = expectedWorkbook.getSheetAt(s);
            Sheet eachActualSheet = actualWorkbook.getSheetAt(s);
            System.out.println("\n\n Expected Sheet : " + eachExpectedSheet.getSheetName() + " \n\n Actual Sheet : " + eachActualSheet.getSheetName() + "\n\n");

            //Create new result sheet with name as Expected sheet name
            Sheet eachNewResultSheet = resultWorkbook.createSheet(eachExpectedSheet.getSheetName());


            //Go Through each row and cell
            System.out.println("\n\nIterating over each Row\n");
            for (int r = eachExpectedSheet.getFirstRowNum(); r <= eachExpectedSheet.getLastRowNum(); r++) {
                String firstRowNumberMatchMessage = eachExpectedSheet.getFirstRowNum() != eachActualSheet.getFirstRowNum() ? "#### FirstRowNum not matched, Expected : " + eachExpectedSheet.getFirstRowNum() + " but Actual : " + eachActualSheet.getFirstRowNum() + "####" : "FirstRowNum matched";
                System.out.println(firstRowNumberMatchMessage);

                String lastRowNumberMatchMessage = eachExpectedSheet.getLastRowNum() != eachActualSheet.getLastRowNum() ? "#### LastRowNum not matched, Expected : " + eachExpectedSheet.getLastRowNum() + " but Actual : " + eachActualSheet.getLastRowNum() + "####" : "LastRowNum matched";
                System.out.println(lastRowNumberMatchMessage);
                Row eachExpectedRow = eachExpectedSheet.getRow(r);
                Row eachActualRow = eachActualSheet.getRow(r);

                Row eachNewResultRow = eachNewResultSheet.createRow(r);


                for (int c = eachExpectedRow.getFirstCellNum(); c < eachExpectedRow.getLastCellNum(); c++) {
                    System.out.println("\n\nIterating over each Column(or) Cell\n");
                    totalCellsCompared = totalCellsCompared + 1;
                    String firstCellNumberMatchMessage = eachExpectedRow.getFirstCellNum() != eachActualRow.getFirstCellNum()
                            ? "#### FirstCellNum not matched, Expected : " + eachExpectedRow.getFirstCellNum() + " but Actual : " + eachActualRow.getFirstCellNum() + "####" : "FirstCellNum matched";
                    System.out.println(firstCellNumberMatchMessage);

                    String lastCellNumberMatchMessage = eachExpectedRow.getLastCellNum() != eachActualRow.getLastCellNum()
                            ? "#### LastCellNum not matched, Expected : " + eachExpectedRow.getLastCellNum() + " but Actual : " + eachActualRow.getLastCellNum() + "####" : "LastCellNum matched";
                    System.out.println(lastCellNumberMatchMessage);

                    Cell eachExpectedCell = eachExpectedRow.getCell(c);
                    Cell eachActualCell = eachActualRow.getCell(c);

                    Cell eachResultCell = eachNewResultRow.createCell(c);

                    CellStyle origStyle = expectedWorkbook.getCellStyleAt(1);
                    CellStyle newStyle = resultWorkbook.createCellStyle();
                    newStyle.cloneStyleFrom(origStyle);

                    //Compare Style
                    if (eachExpectedCell.getRichStringCellValue().toString().equals(eachActualCell.getRichStringCellValue().toString())) {
                        //Match
                        eachResultCell.setCellValue(eachExpectedCell.getRichStringCellValue().toString());
                        eachResultCell.setCellStyle(newStyle);
                        totalCellsMatched = totalCellsMatched + 1;
                        System.out.println("\n\nPlaced Matched Cell\n");
                    } else {
                        Font headerFont = resultWorkbook.createFont();
                        headerFont.setBold(true);
                        headerFont.setFontHeightInPoints((short) 11);
                        headerFont.setColor(IndexedColors.RED.getIndex());
                        //No Match
                        eachResultCell.setCellValue("Expected : " + eachExpectedCell.getRichStringCellValue().toString() +
                                "  Actual :" + eachActualCell.getRichStringCellValue().toString());
                        CellStyle headerCellStyle = resultWorkbook.createCellStyle();
                        headerCellStyle.setFont(headerFont);
                        eachResultCell.setCellStyle(headerCellStyle);
                        totalCellNotMatched = totalCellNotMatched + 1;
                        System.out.println("\n\nPlaced No Matched Cell\n");
                    }
                }
            }
        }
        FileOutputStream fileOut = new FileOutputStream(resultFileWithPath);
        resultWorkbook.write(fileOut);
        fileOut.close();
        System.out.println("totalCellsCompared: " + totalCellsCompared + " totalCellsMatched: " + totalCellsMatched + " totalCellNotMatched: " + totalCellNotMatched);
        //cell.getCellType()
    }
}
