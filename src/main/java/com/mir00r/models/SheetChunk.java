package com.mir00r.models;
import org.apache.poi.ss.usermodel.Sheet;

public class SheetChunk {

  private final Sheet sheet;
  private final int startRow;
  private final int endRow;
  private final String sheetName;

  public SheetChunk(Sheet sheet, int startRow, int endRow, String sheetName) {
    this.sheet = sheet;
    this.startRow = startRow;
    this.endRow = endRow;
    this.sheetName = sheetName;
  }

  public Sheet getSheet() {
    return sheet;
  }

  public int getStartRow() {
    return startRow;
  }

  public int getEndRow() {
    return endRow;
  }

  public String getSheetName() {
    return sheetName;
  }
}
