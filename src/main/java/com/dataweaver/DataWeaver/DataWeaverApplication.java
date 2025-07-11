package com.dataweaver.DataWeaver;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import io.swagger.v3.oas.annotations.OpenAPIDefinition;
import io.swagger.v3.oas.annotations.info.Info;

@OpenAPIDefinition(
  info = @Info(
    title = "DataWeaver application",
    version = "1.0",
    description = "API documentation for DataWeaver application"
  )
)
@SpringBootApplication
public class DataWeaverApplication {

	public static void main(String[] args) {
		SpringApplication.run(DataWeaverApplication.class, args);
	}

}
