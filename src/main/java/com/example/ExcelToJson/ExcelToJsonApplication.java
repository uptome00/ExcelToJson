package com.example.ExcelToJson;

import com.example.ExcelToJson.controller.ExcelController;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.http.ResponseEntity;

import java.io.IOException;

@SpringBootApplication
@Slf4j
public class ExcelToJsonApplication {

	public static void main(String[] args) throws OpenXML4JException, IOException, InterruptedException {
		SpringApplication.run(ExcelToJsonApplication.class, args);

		ExcelController excelController = new ExcelController();


		String path = "D:/Ratchanon/project/spring/ExcelToJson/resource/mockup.xlsx";
		String sheetName = "transaction-Branch";
		String caseNumber = "SIT-ORM-BRH003-002";

//		excelController.allCase(path, sheetName);
		excelController.selectCase(path,sheetName,caseNumber);
	}

}
