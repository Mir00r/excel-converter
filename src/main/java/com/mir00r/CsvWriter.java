package com.mir00r;

import com.mir00r.models.SheetChunk;
import org.apache.poi.ss.usermodel.Row;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;

public class CsvWriter {

  public void writeCsv(SheetChunk sheetChunk, String outputDir) throws IOException {
    File csvFile = new File(outputDir, sheetChunk.getSheetName() + "_chunk.csv");
    try (FileWriter writer = new FileWriter(csvFile)) {
      for (int rowNum = sheetChunk.getStartRow(); rowNum < sheetChunk.getEndRow(); rowNum++) {
        Row row = sheetChunk.getSheet().getRow(rowNum);
        StringBuilder csvLine = new StringBuilder();
        row.forEach(cell -> csvLine.append(cell.toString()).append(","));
        writer.write(csvLine + "\n");
      }
    }
  }
}
