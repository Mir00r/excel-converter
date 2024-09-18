package com.mir00r;

import org.apache.camel.builder.RouteBuilder;
import org.apache.camel.model.dataformat.CsvDataFormat;
import org.springframework.stereotype.Component;


import java.util.List;

@Component
public class ExcelToCSVRoute extends RouteBuilder {

  @Override
  public void configure() throws Exception {

    // Define a route to process Excel
    from("direct:processExcel")
      .routeId("excel-to-csv-route")
      .threads(10) // Use 10 threads for processing
//      .process(new CsvConsumerProcessor(taskQueue, "output-directory"))
      .end();

    // You can also use streaming to avoid loading the entire file
    CsvDataFormat csv = new CsvDataFormat();
    csv.setDelimiter(",");
    csv.setHeader(List.of("column1", "column2")); // Customize columns
  }
}
