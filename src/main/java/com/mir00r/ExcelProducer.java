package com.mir00r;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

public class ExcelProducer implements Runnable {
    private final String excelFilePath;
    private final String csvDirectory;
    private final int chunkSize ;

    public ExcelProducer(String excelFilePath, String csvDirectory, int chunkSize) {
        this.excelFilePath = excelFilePath;
        this.csvDirectory = csvDirectory;
        this.chunkSize = chunkSize;
    }

    @Override
    public void run() {
        ExcelProcessor excelProcessor = new ExcelProcessor(excelFilePath, csvDirectory, chunkSize);
        excelProcessor.processExcel();

//        try (InputStream excelInputStream = new FileInputStream(excelFilePath)) {
//            try (Workbook workbook = new XSSFWorkbook(excelInputStream)) {
//                for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
//                    Sheet sheet = workbook.getSheetAt(sheetIndex);
//                    // Process the sheet and push it to the consumer queue
//                    CsvConsumer.processSheet(sheet);
//                }
//            }
//        } catch (IOException e) {
//            e.printStackTrace();
//        }
    }
}
