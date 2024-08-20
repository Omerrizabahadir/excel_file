package com.example_exelhelper;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

/*
Adımlar:

Numbers Dosyasını Excel Formatında Kaydedin: Numbers uygulamasını kullanarak dosyanızı .xlsx formatında kaydedin.
Java ile Excel Dosyasını İşleyin: Yukarıdaki örnekte olduğu gibi Apache POI kullanarak .xlsx dosyasını okuyun veya yazın.
3. Dosya Formatı Dönüştürme
Numbers dosyalarını Excel formatına dönüştürmek için:

Apple Numbers Uygulamasını Açın.
Dosyayı Açın: Numbers uygulamasında .numbers dosyanızı açın.
Dosya Menüsü > Dışa Aktar > Excel seçeneğini seçin.
Excel Formatında Kaydedin: Dışa aktarılan dosyayı .xlsx formatında kaydedin.
 */


public class ExcelReader {
    public static void main(String[] args) {
        String filePath = "/Users/Macbook/Documents/excel.xlsx";

        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(0); // İlk sayfayı al
            for (Row row : sheet) {
                for (Cell cell : row) {
                    switch (cell.getCellType()) {
                        case STRING:
                            System.out.print(cell.getStringCellValue() + "\t");
                            break;
                        case NUMERIC:
                            if (DateUtil.isCellDateFormatted(cell)) {
                                System.out.print(cell.getDateCellValue() + "\t");
                            } else {
                                System.out.print(cell.getNumericCellValue() + "\t");
                            }
                            break;
                        case BOOLEAN:
                            System.out.print(cell.getBooleanCellValue() + "\t");
                            break;
                        case FORMULA:
                            // Formül hücresinden değer alırken uygun yöntemi kullanın
                            System.out.print(cell.getCellFormula() + "\t");
                            break;
                        default:
                            System.out.print("");
                            break;
                    }
                }
                System.out.println();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}

