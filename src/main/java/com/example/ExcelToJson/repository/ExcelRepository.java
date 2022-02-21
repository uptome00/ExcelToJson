package com.example.ExcelToJson.repository;

import lombok.Data;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;
import org.springframework.stereotype.Repository;

import java.io.FileInputStream;
import java.io.IOException;

@Repository
@Data
public class ExcelRepository extends XSSFWorkbook {

    private FileInputStream fileInputStream;
    private XSSFWorkbook workbook;

    public ExcelRepository(){

    }

    public ExcelRepository(String path) throws IOException {
        this.fileInputStream = new FileInputStream(path);
        this.workbook = new XSSFWorkbook(fileInputStream);
    }

    public XSSFSheet getSheetBySheetIndex(int indexSheet){
        return workbook.getSheetAt(indexSheet);
    }

    public int getSheetIndexBySheetName(String sheetName){
        return workbook.getSheetIndex(sheetName);
    }
}
