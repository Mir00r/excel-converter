package com.mir00r;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.xml.sax.*;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.*;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

public class ExcelFileReaderService {

    public void read () throws Exception {
        String excelFilePath = "c:/Users/sdarmd/Desktop/sdedocs/excel/Employee_Data_1M.xlsx";
        String csvDirectory = "c:/Users/sdarmd/Desktop/sdedocs/csv";
        int chunkSize = 10000; // Adjust the chunk size based on your needs
        String csvSeparator = ",";

//        try (ZipInputStream zipInputStream = new ZipInputStream(new FileInputStream(excelFilePath))) {
//            ZipEntry entry;
//            while ((entry = zipInputStream.getNextEntry()) != null) {
//                if (entry.getName().equals("xl/workbook.xml")) {
//                    processWorkbook(zipInputStream, csvDirectory, csvSeparator);
//                    break;
//                }
//            }
//        } catch (IOException | SAXException e) {
//            throw new RuntimeException(e);
//        }
        FileInputStream excelInputStream = new FileInputStream(excelFilePath);
        XSSFWorkbook workbook = new XSSFWorkbook(excelInputStream);

        // Iterate through each sheet in the Excel file
        for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
            XSSFSheet sheet = workbook.getSheetAt(sheetIndex);

            // Process the sheet and convert it to CSV
            processSheet(sheet, csvDirectory, csvSeparator);
        }

        // Close resources
        workbook.close();
        excelInputStream.close();
    }

    private void processSheet(XSSFSheet sheet, String csvDirectory, String csvSeparator) throws Exception {
        // Create a CSV file for this sheet
        String csvFileName = csvDirectory + sheet.getSheetName() + ".csv";
        FileWriter csvWriter = new FileWriter(csvFileName);

        // Create a data formatter for formatting cell values
        DataFormatter dataFormatter = new DataFormatter();

        // Create an XSSFSheetXMLHandler to process the sheet in a streaming manner
        XSSFSheetXMLHandler xmlHandler = new XSSFSheetXMLHandler(
                null, null, new XSSFSheetXMLHandler.SheetContentsHandler() {
            private boolean isHeaderRow = true;

            @Override
            public void startRow(int rowNum) {
                // Handle the start of a new row
            }

            @Override
            public void endRow(int rowNum) {
                // Handle the end of a row
                try {
                    csvWriter.append('\n');
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }

            @Override
            public void cell(String cellReference, String formattedValue, XSSFComment comment) {
                // Handle cell data
                if (!isHeaderRow) {
                    try {
                        csvWriter.append(formattedValue);
                        csvWriter.append(csvSeparator);
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                }
            }

            @Override
            public void headerFooter(String text, boolean isHeader, String tagName) {
                // Handle header/footer if needed
            }
        }, dataFormatter,false);

//        // Process the sheet using the streaming API
//        XSSFSheetXMLHandlerSheetContentsHandler contentHandler = xmlHandler.getContentsHandler();
//        InputStream sheetInputStream = sheet.getSheetData().getPackagePart().getInputStream();
//        try {
//            ParserFactory.createXMLParser().parse(sheetInputStream, contentHandler);
//        } finally {
//            sheetInputStream.close();
//            csvWriter.close();
//        }
    }

    private static void processWorkbook(InputStream workbookInputStream, String csvDirectory, String csvSeparator) throws IOException, SAXException {
        XSSFWorkbook workbook = new XSSFWorkbook(workbookInputStream);

        for (Sheet sheet : workbook) {
            String sheetName = sheet.getSheetName();
            processSheet(sheet, csvDirectory, sheetName, csvSeparator);
        }
        workbook.close();
    }

    private static void processSheet(Sheet sheet, String csvDirectory, String sheetName, String csvSeparator) throws IOException, SAXException {
        File csvFile = new File(csvDirectory, sheetName + ".csv");
        try (FileOutputStream csvOutputStream = new FileOutputStream(csvFile)) {
            StylesTable styles = ((XSSFWorkbook) sheet.getWorkbook()).getStylesSource();

            XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) sheet.getWorkbook();
            while (iter.hasNext()) {
                try (InputStream sheetInputStream = iter.next()) {
                    CSVSheetContentsHandler contentsHandler = new CSVSheetContentsHandler(csvOutputStream, csvSeparator);
                    processSheetUsingSAX(sheetInputStream, styles, contentsHandler);
                }
            }
        }
    }

    private static void processSheetUsingSAX(InputStream sheetInputStream, StylesTable styles, ContentHandler handler) throws IOException, SAXException {
        XMLReader sheetParser = XMLReaderFactory.createXMLReader();
        XSSFSheetXMLHandler parserHandler = new XSSFSheetXMLHandler(styles, null, (XSSFSheetXMLHandler.SheetContentsHandler) handler, null, true);
        sheetParser.setContentHandler(parserHandler);

        InputSource sheetSource = new InputSource(sheetInputStream);
        sheetSource.setSystemId("");
        sheetParser.parse(sheetSource);
    }

    private static class CSVSheetContentsHandler extends DefaultHandler {
        private final OutputStream csvOutputStream;
        private final String separator;
        private List<String> currentRowData;
        private boolean isCell;

        CSVSheetContentsHandler(OutputStream csvOutputStream, String separator) {
            this.csvOutputStream = csvOutputStream;
            this.separator = separator;
        }

        @Override
        public void startElement(String uri, String localName, String name, Attributes attributes) {
            if (name.equals("c")) {
                isCell = true;
                currentRowData = new ArrayList<>();
            }
        }

        @Override
        public void endElement(String uri, String localName, String name) {
            if (name.equals("c")) {
                if (currentRowData.size() > 0) {
                    try {
                        String rowData = String.join(separator, currentRowData);
                        csvOutputStream.write(rowData.getBytes());
                        csvOutputStream.write("\n".getBytes());
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                }
                isCell = false;
            }
        }

        @Override
        public void characters(char[] ch, int start, int length) {
            if (isCell) {
                currentRowData.add(new String(ch, start, length));
            }
        }
    }
}
