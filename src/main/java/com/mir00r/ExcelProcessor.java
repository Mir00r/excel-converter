package com.mir00r;

import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.xml.sax.*;
import org.xml.sax.helpers.DefaultHandler;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParserFactory;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class ExcelProcessor {
    private final String excelFilePath;
    private final String outputDirectory;
    private final int chunkSize;

    public ExcelProcessor(String excelFilePath, String outputDirectory, int chunkSize) {
        this.excelFilePath = excelFilePath;
        this.outputDirectory = outputDirectory;
        this.chunkSize = chunkSize;
    }

    public void processExcel() {
        try (InputStream excelInputStream = new FileInputStream(excelFilePath)) {
            OPCPackage opcPackage = OPCPackage.open(excelInputStream);
            XSSFReader xssfReader = new XSSFReader(opcPackage);
            SharedStringsTable sharedStringsTable = xssfReader.getSharedStringsTable();

            XMLReader parser = fetchSheetParser(sharedStringsTable);

            Iterator<InputStream> sheets = xssfReader.getSheetsData();
            while (sheets.hasNext()) {
                InputStream sheet = sheets.next();
                InputSource sheetSource = new InputSource(sheet);
                parser.parse(sheetSource);
                sheet.close();
            }

        } catch (IOException | SAXException | ParserConfigurationException | OpenXML4JException e) {
            e.printStackTrace();
        }
    }

    private XMLReader fetchSheetParser(SharedStringsTable sharedStringsTable) throws SAXException, ParserConfigurationException {
        XMLReader parser = SAXParserFactory.newInstance().newSAXParser().getXMLReader();
        ContentHandler handler = new ExcelSheetHandler(sharedStringsTable);
        parser.setContentHandler(handler);
        return parser;
    }

    private static class ExcelSheetHandler extends DefaultHandler {
        private SharedStringsTable sharedStringsTable;
        private StringBuilder cellContent;
        private boolean isCellOpen = false;
        private List<String> rowData = new ArrayList<>();

        public ExcelSheetHandler(SharedStringsTable sharedStringsTable) {
            this.sharedStringsTable = sharedStringsTable;
            cellContent = new StringBuilder();
        }

        @Override
        public void startElement(String uri, String localName, String name, Attributes attributes) {
            if (name.equals("c")) {
                isCellOpen = true;
                cellContent.setLength(0);
            }
        }

        @Override
        public void endElement(String uri, String localName, String name) {
            if (isCellOpen) {
                if (name.equals("v")) {
                    rowData.add(getCellValue(cellContent.toString()));
                }
                if (name.equals("c")) {
                    isCellOpen = false;
                }
            }
        }

        @Override
        public void characters(char[] ch, int start, int length) {
            if (isCellOpen) {
                cellContent.append(ch, start, length);
            }
        }

        @Override
        public void endDocument() {
            // Process the rowData list and convert it to CSV or do other processing as needed
            // rowData will contain the data of a single row from the Excel sheet
            System.out.printf("Row Data: -> "+rowData.get(0) + "\n");
        }

        private String getCellValue(String cellValue) {
            // Use sharedStringsTable to convert the cellValue if it's a shared string
            System.out.printf("getCellValue method called: -> "+cellValue +"\n");
            int idx = Integer.parseInt(cellValue);
            if (sharedStringsTable.getItemAt(idx) != null) {
                return sharedStringsTable.getItemAt(idx).getString();
            }
            return cellValue;
        }
    }
}