package com.mir00r.routes;

import com.mir00r.CsvWriter;
import com.mir00r.models.SheetChunk;
import org.apache.camel.builder.RouteBuilder;
import org.springframework.stereotype.Component;

@Component
public class CsvRoute extends RouteBuilder {

  public static final String DIRECT_PROCESS_EXCEL = "direct:processSheet";

  @Override
  public void configure() throws Exception {
    from(DIRECT_PROCESS_EXCEL)
      .routeId("processSheetRoute")
      .log("Processing Excel sheet data Body: ${body}")
      .process(exchange -> {
        SheetChunk sheetChunk = exchange.getIn().getBody(SheetChunk.class);
        String outputDir = "output/csv";
        new CsvWriter().writeCsv(sheetChunk, outputDir);
      })
      .to("file://output/csv")  // Or any other endpoint
      .log("Excel sheet processing completed");
  }
}
