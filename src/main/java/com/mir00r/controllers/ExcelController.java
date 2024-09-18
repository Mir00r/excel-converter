package com.mir00r.controllers;

import com.mir00r.ZipCompressor;
import com.mir00r.models.SheetChunk;
import com.mir00r.services.ExcelService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.IOException;
import java.util.concurrent.ArrayBlockingQueue;
import java.util.concurrent.BlockingQueue;

@RestController
@RequestMapping("/api/v1/excel")
public class ExcelController {

  private final ExcelService excelService;

  @Autowired
  public ExcelController(ExcelService excelService) {
    this.excelService = excelService;
  }

  @PostMapping("/process")
  public String processExcel(@RequestParam("file") MultipartFile file) throws IOException, InterruptedException {
    File outputDirectory = new File("output/csv");
    if (!outputDirectory.exists()) {
      outputDirectory.mkdirs();
    }

    BlockingQueue<SheetChunk> taskQueue = new ArrayBlockingQueue<>(10); // Adjust the size as needed
    excelService.processExcel(file, taskQueue);

    // Compress the CSV files
    new ZipCompressor().compressCSVFiles(outputDirectory);

    return "Processing started, compressed CSV files will be available soon.";
  }
}
