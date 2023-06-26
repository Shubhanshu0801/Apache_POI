package com.apachepoi;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;

public class ReadFromExcelFile {
    public static void main(String[] args) throws Exception {
        File file = new File("C:\\Users\\Ishani\\IdeaProjects\\Apache_POI\\Excel_files\\Cricket_team.xlsx");
        FileInputStream fis = new FileInputStream(file);
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        XSSFSheet sheet = workbook.getSheet("World_XI");
        int rowCount = sheet.getPhysicalNumberOfRows();
        for(int i=0; i<rowCount; i++) {
            XSSFRow row = sheet.getRow(i);
            int columnCount = row.getPhysicalNumberOfCells();
            for(int j=0; j<columnCount; j++) {
                XSSFCell cell = row.getCell(j);
                String cellValue = getCellValue(cell);
                System.out.println(cellValue);
            }
            System.out.println();
        }
        workbook.close();
        fis.close();
    }
    public static String getCellValue(XSSFCell cell) {
        switch(cell.getCellType()) {
            case NUMERIC:
                return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case STRING:
                return cell.getStringCellValue();
            default:
                return String.valueOf(cell.getNumericCellValue());
        }
    }

}
