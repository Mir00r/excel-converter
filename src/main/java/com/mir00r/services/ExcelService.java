package com.mir00r.services;

import com.mir00r.models.SheetChunk;
import com.mir00r.routes.CsvRoute;
import org.apache.camel.ProducerTemplate;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.scheduling.annotation.Async;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.util.concurrent.BlockingQueue;

@Service
public class ExcelService {

  private final ProducerTemplate producerTemplate;

  @Autowired
  public ExcelService(ProducerTemplate producerTemplate) {
    this.producerTemplate = producerTemplate;
  }

  @Async
  public void processExcel(MultipartFile excelFile, BlockingQueue<SheetChunk> taskQueue) throws
    IOException {
    try (InputStream fis = excelFile.getInputStream();
      Workbook workbook = new XSSFWorkbook(fis)) {

      // Iterate over sheets
      for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
        Sheet sheet = workbook.getSheetAt(i);
        int rows = sheet.getPhysicalNumberOfRows();

        // Iterate in chunks (e.g., 1000 rows per chunk)
        int chunkSize = 1000;
        for (int startRow = 0; startRow < rows; startRow += chunkSize) {
          int endRow = Math.min(startRow + chunkSize, rows);
          // Create and put chunk in taskQueue
          SheetChunk chunk = new SheetChunk(sheet, startRow, endRow, "sheet_" + i);
          taskQueue.put(chunk);
          producerTemplate.sendBody(CsvRoute.DIRECT_PROCESS_EXCEL, chunk); // Send to Camel route
        }
      }
    } catch (InterruptedException e) {
      Thread.currentThread().interrupt();
      throw new RuntimeException("Processing interrupted", e);
    }
  }
}
