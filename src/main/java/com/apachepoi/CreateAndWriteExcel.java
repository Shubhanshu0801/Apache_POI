package com.apachepoi;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class CreateAndWriteExcel {
    public static void main(String[] args) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("World_XI");
        sheet.createRow(0);
        sheet.getRow(0).createCell(0).setCellValue("Name");
        sheet.getRow(0).createCell(1).setCellValue("Country");
        sheet.getRow(0).createCell(2).setCellValue("Roles");
        sheet.getRow(0).createCell(3).setCellValue("Batting-average");
        sheet.getRow(0).createCell(4).setCellValue("Bowling-average");

        sheet.createRow(1);
        sheet.getRow(1).createCell(0).setCellValue("Babar Azam");
        sheet.getRow(1).createCell(1).setCellValue("Pakistan");
        sheet.getRow(1).createCell(2).setCellValue("Batsman");
        sheet.getRow(1).createCell(3).setCellValue(59.17);
        sheet.getRow(1).createCell(4).setCellValue(00.00);

        sheet.createRow(2);
        sheet.getRow(2).createCell(0).setCellValue("David Warner");
        sheet.getRow(2).createCell(1).setCellValue("Australia");
        sheet.getRow(2).createCell(2).setCellValue("Batsman");
        sheet.getRow(2).createCell(3).setCellValue(45.00);
        sheet.getRow(2).createCell(4).setCellValue(00.00);

        sheet.createRow(3);
        sheet.getRow(3).createCell(0).setCellValue("Steve Smith");
        sheet.getRow(3).createCell(1).setCellValue("Australia");
        sheet.getRow(3).createCell(2).setCellValue("Batsman");
        sheet.getRow(3).createCell(3).setCellValue(44.49);
        sheet.getRow(3).createCell(4).setCellValue(34.67);

        sheet.createRow(4);
        sheet.getRow(4).createCell(0).setCellValue("AB Devilliers");
        sheet.getRow(4).createCell(1).setCellValue("South Africa");
        sheet.getRow(4).createCell(2).setCellValue("batsman");
        sheet.getRow(4).createCell(3).setCellValue(53.50);
        sheet.getRow(4).createCell(4).setCellValue(28.85);

        sheet.createRow(5);
        sheet.getRow(5).createCell(0).setCellValue("Jos Butler");
        sheet.getRow(5).createCell(1).setCellValue("England");
        sheet.getRow(5).createCell(2).setCellValue("Batsman");
        sheet.getRow(5).createCell(3).setCellValue(41.61);
        sheet.getRow(5).createCell(4).setCellValue(00.00);

        sheet.createRow(6);
        sheet.getRow(6).createCell(0).setCellValue("Yuvraj Singh");
        sheet.getRow(6).createCell(1).setCellValue("India");
        sheet.getRow(6).createCell(2).setCellValue("Batsman/Bowler");
        sheet.getRow(6).createCell(3).setCellValue(36.55);
        sheet.getRow(6).createCell(4).setCellValue(38.68);

        sheet.createRow(7);
        sheet.getRow(7).createCell(0).setCellValue("MS Dhoni(C/WK)");
        sheet.getRow(7).createCell(1).setCellValue("India");
        sheet.getRow(7).createCell(2).setCellValue("Batsman");
        sheet.getRow(7).createCell(3).setCellValue(50.53);
        sheet.getRow(7).createCell(4).setCellValue(31.00);

        sheet.createRow(8);
        sheet.getRow(8).createCell(0).setCellValue("Shaheen Afridi");
        sheet.getRow(8).createCell(1).setCellValue("Pakistan");
        sheet.getRow(8).createCell(2).setCellValue("Bowler/Batsman");
        sheet.getRow(8).createCell(3).setCellValue(17.85);
        sheet.getRow(8).createCell(4).setCellValue(23.94);

        sheet.createRow(9);
        sheet.getRow(9).createCell(0).setCellValue("Haris Rauf");
        sheet.getRow(9).createCell(1).setCellValue("Pakistan");
        sheet.getRow(9).createCell(2).setCellValue("Bowler");
        sheet.getRow(9).createCell(3).setCellValue(2.75);
        sheet.getRow(9).createCell(4).setCellValue(28.28);

        sheet.createRow(10);
        sheet.getRow(10).createCell(0).setCellValue("Josh Hazlewood");
        sheet.getRow(10).createCell(1).setCellValue("Australia");
        sheet.getRow(10).createCell(2).setCellValue("Bowler");
        sheet.getRow(10).createCell(3).setCellValue(20.25);
        sheet.getRow(10).createCell(4).setCellValue(25.82);

        sheet.createRow(11);
        sheet.getRow(11).createCell(0).setCellValue("James Anderson");
        sheet.getRow(11).createCell(1).setCellValue("England");
        sheet.getRow(11).createCell(2).setCellValue("Bowler");
        sheet.getRow(11).createCell(3).setCellValue(7.58);
        sheet.getRow(11).createCell(4).setCellValue(29.22);

        File file = new File("C:\\Users\\Ishani\\IdeaProjects\\Apache_POI\\Excel_files\\Cricket_team.xlsx");
        FileOutputStream fos = new FileOutputStream(file);
        workbook.write(fos);
        fos.close();
    }
}
