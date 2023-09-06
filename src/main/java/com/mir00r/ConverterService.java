package com.mir00r;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class ConverterService {

    public void convert() throws IOException {
        String excelFilePath = "c:/Users/sdarmd/Desktop/sdedocs/excel/Employee_Data_1M.xlsx";
        String csvDirectory = "c:/Users/sdarmd/Desktop/sdedocs/csv";
        int chunkSize = 10000; // Adjust the chunk size based on your needs
        String csvSeparator = ",";

        try (InputStream excelInputStream = new FileInputStream(excelFilePath)) {
            try (Workbook workbook = new XSSFWorkbook(excelInputStream)) {
                // Iterate through each sheet in the Excel file
                for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
                    Sheet sheet = workbook.getSheetAt(sheetIndex);

                    // Process the sheet and convert it to CSV
                    this.processSheet(sheet, csvDirectory, chunkSize, csvSeparator);
                }
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        }
    }

    private void processSheet(Sheet sheet, String csvDirectory, int chunkSize, String csvSeparator) throws IOException {
        int numRows = sheet.getPhysicalNumberOfRows();
        List<String> header = getHeader(sheet.getRow(0));

        // Iterate through the rows and split into chunks
        for (int chunkStart = 1; chunkStart < numRows; chunkStart += chunkSize) {
            int chunkEnd = Math.min(chunkStart + chunkSize, numRows);

            List<List<String>> dataChunk = new ArrayList<>();
            dataChunk.add(header); // Add the header to each chunk

            for (int rowNum = chunkStart; rowNum < chunkEnd; rowNum++) {
                Row row = sheet.getRow(rowNum);
                List<String> rowData = getRowData(row);
                dataChunk.add(rowData);
            }

            // Convert the chunk to CSV and save it
            saveChunkToCsv(dataChunk, csvDirectory, sheet.getSheetName(), chunkStart / chunkSize, csvSeparator);
        }
    }

    private List<String> getHeader(Row headerRow) {
        List<String> header = new ArrayList<>();
        for (Cell cell : headerRow) {
            header.add(cell.getStringCellValue());
        }
        return header;
    }

    private List<String> getRowData(Row row) {
        List<String> rowData = new ArrayList<>();
        for (Cell cell : row) {
            rowData.add(cell.getStringCellValue());
        }
        return rowData;
    }

    private void saveChunkToCsv(List<List<String>> data, String csvDirectory, String sheetName, int chunkIndex, String separator) throws IOException {
        String csvFileName = csvDirectory + "output_" + sheetName + "_part" + (chunkIndex + 1) + ".csv";

        try (Writer writer = new FileWriter(csvFileName)) {
            for (List<String> rowData : data) {
                String line = String.join(separator, rowData);
                writer.write(line);
                writer.write("\n");
            }
        }
        System.out.println("Saved CSV chunk: " + csvFileName);
    }
}
