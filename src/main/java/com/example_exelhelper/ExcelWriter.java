package com.example_exelhelper;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelWriter {

    public static void main(String[] args) {
        // Çıktı dosyasının yolu
        String outputFilePath = "/Users/Macbook/Documents/output.xlsx";

        // Workbook ve Sheet oluştur
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Students");

            // Hücre stili oluşturma
            CellStyle headerStyle = workbook.createCellStyle();
            CellStyle cellStyle = workbook.createCellStyle();

            // Font oluşturma
            Font headerFont = workbook.createFont();
            headerFont.setBold(true);
            headerFont.setColor(IndexedColors.WHITE.getIndex());
            headerStyle.setFont(headerFont);
            headerStyle.setFillForegroundColor(IndexedColors.BLUE.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            Font cellFont = workbook.createFont();
            cellFont.setBold(false);
            cellStyle.setFont(cellFont);
            cellStyle.setBorderBottom(BorderStyle.THIN);
            cellStyle.setBorderTop(BorderStyle.THIN);
            cellStyle.setBorderLeft(BorderStyle.THIN);
            cellStyle.setBorderRight(BorderStyle.THIN);
            cellStyle.setAlignment(HorizontalAlignment.CENTER);

            // Başlık satırı oluşturma
            //tablonun başlıkları
            Row headerRow = sheet.createRow(0);
            Cell headerCell1 = headerRow.createCell(0);
            headerCell1.setCellValue("Name");
            headerCell1.setCellStyle(headerStyle);

            Cell headerCell2 = headerRow.createCell(1);
            headerCell2.setCellValue("Age");
            headerCell2.setCellStyle(headerStyle);

            Cell headerCell3 = headerRow.createCell(2);
            headerCell3.setCellValue("Number");
            headerCell3.setCellStyle(headerStyle);

            Cell headerCell4 = headerRow.createCell(3);
            headerCell4.setCellValue("Gender");
            headerCell4.setCellStyle(headerStyle);


            //tablodaki datalar
            // Veri satırı oluşturma
            //1.sütunun ilk satırı.1. öğrenci
            Row row1 = sheet.createRow(1);          //1.row(satır)
            Cell cell1 = row1.createCell(0);        //1.row 'un 0. cell'i(hücresi)
            cell1.setCellValue("Alice");
            cell1.setCellStyle(cellStyle);

            Cell cell2 = row1.createCell(1);        //1.row 'un 1. cell'i(hücresi)
            cell2.setCellValue(10);
            cell2.setCellStyle(cellStyle);

            Cell cell3 = row1.createCell(2);        //1.row 'un 2. cell'i(hücresi)
            cell3.setCellValue(1);
            cell3.setCellStyle(cellStyle);

            Cell cell4 = row1.createCell(3);        //1.row 'un 3. cell'i(hücresi)
            cell4.setCellValue("female");
            cell4.setCellStyle(cellStyle);

            //1. sütun 2.satır. yani 2. öğrenci için

            Row row2 = sheet.createRow(2);              //2. row
            Cell cell5 = row2.createCell(0);            //2.row 'un 0. cell'i(hücresi)
            cell5.setCellValue("Bob");
            cell5.setCellStyle(cellStyle);

            Cell cell6 = row2.createCell(1);
            cell6.setCellValue(11);
            cell6.setCellStyle(cellStyle);

            Cell cell7 = row2.createCell(2);
            cell7.setCellValue(2);
            cell7.setCellStyle(cellStyle);

            Cell cell8 = row2.createCell(3);
            cell8.setCellValue("male");
            cell8.setCellStyle(cellStyle);

            //1.sütun 3. satır.3. ÖĞRENCİ
            Row row3 = sheet.createRow(3);               //3. row
            Cell cell9 = row3.createCell(0);
            cell9.setCellValue("Cathy");
            cell9.setCellStyle(cellStyle);

            Cell cell10 = row3.createCell(1);
            cell10.setCellValue(12);
            cell10.setCellStyle(cellStyle);

            Cell cell11 = row3.createCell(2);
            cell11.setCellValue(3);
            cell11.setCellStyle(cellStyle);

            Cell cell12 = row3.createCell(3);
            cell12.setCellValue("female");
            cell12.setCellStyle(cellStyle);


            // Dosyayı yazma
            try (FileOutputStream fos = new FileOutputStream(outputFilePath)) {
                workbook.write(fos);
            }

            System.out.println("Excel dosyası stil ile başarıyla oluşturuldu.");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
