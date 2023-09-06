package com.mir00r;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.BlockingQueue;
import java.util.concurrent.LinkedBlockingQueue;

public class CsvConsumer implements Runnable {

    private static final int CHUNK_SIZE = 1000;
    private static final String CSV_SEPARATOR = ",";

    private final String csvDirectory;
    private static final BlockingQueue<List<String>> csvQueue = new LinkedBlockingQueue<>();

    public CsvConsumer(String csvDirectory) {
        this.csvDirectory = csvDirectory;
    }

    @Override
    public void run() {
        while (true) {
            try {
                List<String> csvData = csvQueue.take();
                if (csvData == null || csvData.isEmpty()) {
                    break;
                }

                String sheetName = csvData.remove(0);
                int chunkIndex = Integer.parseInt(csvData.remove(0));
                saveChunkToCsv(sheetName, chunkIndex, csvData);
            } catch (InterruptedException e) {
                Thread.currentThread().interrupt();
                break;
            }
        }
    }

    public static void processSheet(Sheet sheet) {
        int numRows = sheet.getPhysicalNumberOfRows();
        List<String> header = getHeader(sheet.getRow(0));

        for (int chunkStart = 1; chunkStart < numRows; chunkStart += CHUNK_SIZE) {
            int chunkEnd = Math.min(chunkStart + CHUNK_SIZE, numRows);

            List<List<String>> dataChunk = new ArrayList<>();
            dataChunk.add(header);

            for (int rowNum = chunkStart; rowNum < chunkEnd; rowNum++) {
                Row row = sheet.getRow(rowNum);
                List<String> rowData = getRowData(row);
                dataChunk.add(rowData);
            }

            try {
                String sheetName = sheet.getSheetName();
                List<String> csvChunk = new ArrayList<>();
                csvChunk.add(sheetName);
                csvChunk.add(String.valueOf(chunkStart / CHUNK_SIZE));
                for (List<String> row : dataChunk) {
                    csvChunk.addAll(row);
                }
                csvQueue.put(csvChunk);
            } catch (InterruptedException e) {
                Thread.currentThread().interrupt();
            }
        }
    }

    private static List<String> getHeader(Row headerRow) {
        List<String> header = new ArrayList<>();
        for (Cell cell : headerRow) {
            header.add(AppUtil.getCellValueAsString(cell));
        }
        return header;
    }

    private static List<String> getRowData(Row row) {
        List<String> rowData = new ArrayList<>();
        for (Cell cell : row) {
            rowData.add(AppUtil.getCellValueAsString(cell));
        }
        return rowData;
    }

    private void saveChunkToCsv(String sheetName, int chunkIndex, List<String> data) {
        String csvFileName = csvDirectory + "output_" + sheetName + "_part" + (chunkIndex + 1) + ".csv";
        try (FileWriter writer = new FileWriter(csvFileName)) {
            for (String row : data) {
                writer.write(row);
                writer.write("\n");
            }
            System.out.println("Saved CSV chunk: " + csvFileName);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
