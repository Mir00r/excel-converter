package com.mir00r;

import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

public class ConverterWithThreadService {

    public void convert() {
        String excelFilePath = "c:/Users/sdarmd/Desktop/sdedocs/excel/Employee_Data_1M.xlsx";
//        String excelFilePath = "c:/Users/sdarmd/Desktop/sdedocs/excel/STIP_Approved_&_Non-Approved_transactions.xlsx";
        String csvDirectory = "c:/Users/sdarmd/Desktop/sdedocs/csv/";
        int numProducerThreads = 2;
        int numConsumerThreads = 4;

        ExecutorService producerExecutor = Executors.newFixedThreadPool(numProducerThreads);
        ExecutorService consumerExecutor = Executors.newFixedThreadPool(numConsumerThreads);

        try {
            // Initialize and start producer threads
            for (int i = 0; i < numProducerThreads; i++) {
                producerExecutor.execute(new ExcelProducer(excelFilePath, csvDirectory, 100000));
            }

            // Initialize and start consumer threads
            for (int i = 0; i < numConsumerThreads; i++) {
                consumerExecutor.execute(new CsvConsumer(csvDirectory));
            }
        } finally {
            // Shutdown the thread pools when done
            producerExecutor.shutdown();
            consumerExecutor.shutdown();
        }
    }
}
