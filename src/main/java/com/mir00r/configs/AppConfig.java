package com.mir00r.configs;

import org.apache.camel.CamelConfiguration;
import org.apache.camel.CamelContext;
import org.apache.camel.ProducerTemplate;
import org.apache.camel.impl.DefaultCamelContext;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

import java.io.InputStream;
import java.util.concurrent.ArrayBlockingQueue;
import java.util.concurrent.BlockingQueue;

@Configuration
public class AppConfig implements CamelConfiguration {

  @Bean
  public BlockingQueue<InputStream> taskQueue() {
    // Define the taskQueue with a fixed capacity (e.g., 100 items)
    return new ArrayBlockingQueue<>(100);  // The capacity can be adjusted as needed
  }

  @Bean
  public CamelContext camelContext() {
    CamelContext camelContext = new DefaultCamelContext();
    camelContext.start();  // Ensure it's started
    return camelContext;
  }

  @Bean
  public ProducerTemplate producerTemplate(CamelContext camelContext) {
    return camelContext.createProducerTemplate();
  }
}
