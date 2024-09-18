package com.mir00r;

import org.apache.camel.CamelContext;
import org.apache.camel.ProducerTemplate;
import org.apache.camel.impl.DefaultCamelContext;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.*;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.ArrayBlockingQueue;
import java.util.concurrent.BlockingQueue;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

@SpringBootApplication
public class Main {

    private static final int CHUNK_SIZE = 10000; // Adjust the chunk size as needed
    private static final String CSV_SEPARATOR = ","; // Change to your desired CSV separator

    public static void main(String[] args) throws Exception {

        SpringApplication.run(Main.class, args);

        // Initialize Camel context
//        CamelContext camelContext = new DefaultCamelContext();
//        ProducerTemplate producerTemplate = camelContext.createProducerTemplate();
//        ExecutorService producerExecutor = Executors.newFixedThreadPool(2);
//        ExecutorService consumerExecutor = Executors.newFixedThreadPool(5);
//
//        BlockingQueue<InputStream> taskQueue = new ArrayBlockingQueue<>(1000); // Queue for Excel data streams
//
//        // Start producer and consumer threads
//        producerExecutor.execute(new ExcelProducer("path/to/excel.xlsx", taskQueue, producerTemplate));
//        consumerExecutor.execute(new CsvConsumerProcessor(taskQueue, "path/to/output"));
//
//        // Shutdown executors after processing
//        producerExecutor.shutdown();
//        consumerExecutor.shutdown();




//        new ConverterService().convert();
//        new ExcelFileReaderService().read();
//        new ConverterWithThreadService().convert();

//        String excelFilePath = "c:/Users/sdarmd/Desktop/sdedocs/excel/Employee_Data_1M.xlsx"; // Provide the path to your Excel file
//        String outputDirectory = "c:/Users/sdarmd/Desktop/sdedocs/csv"; // Provide the path to your output directory
//
//        try (InputStream excelInputStream = new FileInputStream(excelFilePath);
//             Workbook workbook = new XSSFWorkbook(excelInputStream)) {
//
//            ExecutorService executorService = Executors.newFixedThreadPool(4); // Adjust the number of threads as needed
//
//            for (Sheet sheet : workbook) {
//                int sheetIndex = workbook.getSheetIndex(sheet);
//                executorService.execute(() -> processSheet(sheet, sheetIndex, outputDirectory));
//            }
//
//            executorService.shutdown();
//        } catch (IOException e) {
//            e.printStackTrace();
//        }
    }

    private static void processSheet(Sheet sheet, int sheetIndex, String outputDirectory) {
        int rowCount = sheet.getPhysicalNumberOfRows();

        for (int i = 0; i < rowCount; i += CHUNK_SIZE) {
            int startRow = i;
            int endRow = Math.min(i + CHUNK_SIZE, rowCount);

            List<String[]> dataRows = new ArrayList<>();

            for (int rowNum = startRow; rowNum < endRow; rowNum++) {
                Row row = sheet.getRow(rowNum);
                if (row != null) {
                    int cellCount = row.getPhysicalNumberOfCells();
                    String[] data = new String[cellCount];

                    for (int cellNum = 0; cellNum < cellCount; cellNum++) {
                        Cell cell = row.getCell(cellNum, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                        if (cell != null) {
                            switch (cell.getCellType()) {
                                case STRING:
                                    data[cellNum] = cell.getStringCellValue();
                                    break;
                                case NUMERIC:
                                    data[cellNum] = String.valueOf(cell.getNumericCellValue());
                                    break;
                                case BOOLEAN:
                                    data[cellNum] = String.valueOf(cell.getBooleanCellValue());
                                    break;
                                default:
                                    data[cellNum] = "";
                                    break;
                            }
                        } else {
                            data[cellNum] = "";
                        }
                    }
                    dataRows.add(data);
                }
            }

            writeCsv(dataRows, sheetIndex, i / CHUNK_SIZE, outputDirectory);
        }
    }

    private static void writeCsv(List<String[]> dataRows, int sheetIndex, int chunkIndex, String outputDirectory) {
        String csvFilePath = outputDirectory + "/output_sheet" + sheetIndex + "_part" + chunkIndex + ".csv";

        try (BufferedWriter writer = new BufferedWriter(new FileWriter(csvFilePath))) {
            for (String[] data : dataRows) {
                for (int i = 0; i < data.length; i++) {
                    writer.write(data[i]);
                    if (i < data.length - 1) {
                        writer.write(CSV_SEPARATOR);
                    }
                }
                writer.newLine();
            }
            System.out.println("CSV file created: " + csvFilePath);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
