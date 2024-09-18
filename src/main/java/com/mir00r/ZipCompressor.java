package com.mir00r;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class ZipCompressor {

  public void compressCSVFiles(File outputDirectory) throws IOException {
    File[] csvFiles = outputDirectory.listFiles((dir, name) -> name.endsWith(".csv"));
    try (ZipOutputStream zipOut = new ZipOutputStream(new FileOutputStream(new File(outputDirectory, "output.zip")))) {
      for (File csvFile : csvFiles) {
        try (FileInputStream fis = new FileInputStream(csvFile)) {
          ZipEntry zipEntry = new ZipEntry(csvFile.getName());
          zipOut.putNextEntry(zipEntry);
          byte[] buffer = new byte[1024];
          int len;
          while ((len = fis.read(buffer)) > 0) {
            zipOut.write(buffer, 0, len);
          }
          zipOut.closeEntry();
        }
      }
    }
  }
}
