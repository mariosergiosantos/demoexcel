package com.example.cache.demo;

import com.example.cache.demo.service.GenerateExcel;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class DemoApplication {


	private GenerateExcel generateExcel;

	public static void main(String[] args) throws Exception {
//		GenerateExcel generateExcel = new GenerateExcel();
//
//		generateExcel.process(generateExcel.mock());

		SpringApplication.run(DemoApplication.class, args);
	}

}
