package com.mir00r;

import org.apache.camel.ProducerTemplate;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.concurrent.BlockingQueue;

public class ExcelProducer implements Runnable {

    private final String excelFilePath;
    private final BlockingQueue<InputStream> taskQueue;
    private final ProducerTemplate producerTemplate;

    public ExcelProducer(String excelFilePath, BlockingQueue<InputStream> taskQueue, ProducerTemplate producerTemplate) {
        this.excelFilePath = excelFilePath;
        this.taskQueue = taskQueue;
        this.producerTemplate = producerTemplate;
    }

    @Override
    public void run() {
        try {
            // Step 1: Open the Excel file as an OPCPackage
            OPCPackage pkg = OPCPackage.open(new File(excelFilePath));

            // Step 2: Create the XSSFReader using the OPCPackage
            XSSFReader reader = new XSSFReader(pkg);
            XSSFReader.SheetIterator sheets = (XSSFReader.SheetIterator) reader.getSheetsData();
            while (sheets.hasNext()) {
                try (InputStream sheetInputStream = sheets.next()) {
                    taskQueue.put(sheetInputStream); // Add the sheet's InputStream to the queue
                    producerTemplate.sendBody("direct:processExcel", sheetInputStream); // Send to consumer route
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }


//    private final String excelFilePath;
//    private final String csvDirectory;
//    private final int chunkSize ;
//
//    public ExcelProducer(String excelFilePath, String csvDirectory, int chunkSize) {
//        this.excelFilePath = excelFilePath;
//        this.csvDirectory = csvDirectory;
//        this.chunkSize = chunkSize;
//    }
//
//    @Override
//    public void run() {
//        ExcelProcessor excelProcessor = new ExcelProcessor(excelFilePath, csvDirectory, chunkSize);
//        excelProcessor.processExcel();
//
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
//    }
}
