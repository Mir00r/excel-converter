package com.mir00r;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRst;
import org.xml.sax.*;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.*;

public class ConverterWithStreamingService {
//
//    private SharedStringsTable sharedStrings;
//    private boolean hasMoreRows;
//
//    public void convert() {
//        String excelFilePath = "c:/Users/sdarmd/Desktop/sdedocs/excel/Employee_Data_1M.xlsx";
//        String csvDirectory = "c:/Users/sdarmd/Desktop/sdedocs/csv";
//        int chunkSize = 10000; // Adjust the chunk size based on your needs
//        String csvSeparator = ",";
//
//        try (InputStream excelInputStream = new FileInputStream(excelFilePath)) {
//            sharedStrings = new SharedStringsTable();
//            processExcel(excelInputStream, csvDirectory, chunkSize, csvSeparator);
//        } catch (IOException | SAXException | XmlException | OpenXML4JException e) {
//            e.printStackTrace();
//        }
//    }
//
//    private void processExcel(InputStream excelInputStream, String csvDirectory, int chunkSize, String csvSeparator) throws IOException, SAXException, XmlException, OpenXML4JException {
//        OPCPackage opcPackage = OPCPackage.open(excelInputStream);
//        XSSFReader xssfReader = new XSSFReader(opcPackage);
//        XSSFReader.SheetIterator sheets = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
//
//        while (sheets.hasNext()) {
//            try (InputStream sheetInputStream = sheets.next()) {
//                processSheet(sheetInputStream, csvDirectory, chunkSize, csvSeparator);
//            }
//        }
//
//        opcPackage.close(); // Close the OPCPackage to release resources
//    }
//
//    private void processSheet(InputStream sheetInputStream, String csvDirectory, int chunkSize, String csvSeparator) throws IOException, SAXException {
//        // Set up SAX parser and handlers for reading the sheet XML
//        XMLReader sheetParser = XMLReaderFactory.createXMLReader();
//        final String[] header = new String[1];
//        final String[][] chunkData = new String[chunkSize + 1][]; // Extra 1 for the header
//        chunkData[0] = header;
//        final int[] rowCount = {0};
//        final int[] chunkRowCount = {1};
//
//        ContentHandler handler = new XSSFSheetXMLHandler(null, null, new DefaultHandler() {
//            private StringBuilder cellValue = new StringBuilder();
//            private int currentColumn = -1;
//
//            @Override
//            public void startElement(String uri, String localName, String name, Attributes attributes) {
//                if (name.equals("c")) {
//                    String cellType = attributes.getValue("t");
//                    if (cellType != null && cellType.equals("s")) {
//                        // Cell type is shared string (rich text)
//                        currentColumn = -1;
//                    } else {
//                        // Cell type is not shared string
//                        currentColumn++;
//                    }
//                    cellValue.setLength(0);
//                }
//            }
//
//            @Override
//            public void characters(char[] ch, int start, int length) {
//                cellValue.append(ch, start, length);
//            }
//
//            @Override
//            public void endElement(String uri, String localName, String name) {
//                switch (name) {
//                    case "v":
//                        // Cell value
//                        if (currentColumn == -1) {
//                            int sharedStringIndex = Integer.parseInt(cellValue.toString());
//                            XSSFRichTextString richText = new XSSFRichTextString((CTRst) sharedStrings.getItemAt(sharedStringIndex));
//                            chunkData[chunkRowCount[0]][currentColumn] = richText.toString();
//                        } else {
//                            chunkData[chunkRowCount][currentColumn] = cellValue.toString();
//                        }
//                        break;
//                    case "row":
//                        // End of a row
//                        rowCount[0]++;
//                        chunkRowCount[0]++;
//
//                        if (chunkRowCount[0] > chunkSize) {
//                            // Convert and save the chunk to CSV
//                            saveChunkToCsv(chunkData, csvDirectory, header[0], rowCount[0] / chunkSize, csvSeparator);
//                            chunkRowCount[0] = 1; // Reset for the next chunk
//                        }
//                        break;
//                    case "sheetData":
//                        // End of the sheet
//                        hasMoreRows = false;
//                        break;
//                }
//            }
//        }, false);
//
//        sheetParser.setContentHandler(handler);
//
//        InputSource sheetSource = new InputSource(sheetInputStream);
//        sheetParser.parse(sheetSource);
//    }
//
//    private void saveChunkToCsv(String[][] data, String csvDirectory, String sheetName, int chunkIndex, String separator) {
//        String csvFileName = csvDirectory + "output_" + sheetName + "_part" + chunkIndex + ".csv";
//
//        try (Writer writer = new FileWriter(csvFileName)) {
//            for (String[] rowData : data) {
//                String line = String.join(separator, rowData);
//                writer.write(line);
//                writer.write("\n");
//            }
//        } catch (IOException e) {
//            e.printStackTrace();
//        }
//
//        System.out.println("Saved CSV chunk: " + csvFileName);
//    }
}
