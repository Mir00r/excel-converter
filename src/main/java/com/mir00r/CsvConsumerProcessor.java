package com.mir00r;

import org.apache.camel.Exchange;
import org.apache.camel.Processor;

import java.io.BufferedReader;
import java.io.FileWriter;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.PrintWriter;
import java.util.concurrent.BlockingQueue;

public class CsvConsumerProcessor implements Processor {

  private final BlockingQueue<InputStream> taskQueue;
  private final String outputDirectory;

  public CsvConsumerProcessor(BlockingQueue<InputStream> taskQueue, String outputDirectory) {
    this.taskQueue = taskQueue;
    this.outputDirectory = outputDirectory;
  }

  @Override
  public void process(Exchange exchange) throws Exception {
    InputStream sheetStream = taskQueue.take();
    try (BufferedReader reader = new BufferedReader(new InputStreamReader(sheetStream));
      PrintWriter writer = new PrintWriter(new FileWriter(outputDirectory + "/output.csv", true))) {

      String line;
      while ((line = reader.readLine()) != null) {
        // Process each line and write to CSV
        writer.println(line);
      }
    }
  }
}
