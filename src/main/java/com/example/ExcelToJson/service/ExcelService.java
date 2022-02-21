package com.example.ExcelToJson.service;

import com.example.ExcelToJson.repository.ExcelRepository;
import lombok.Data;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicReference;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

@Service
@Data
@Slf4j
public class ExcelService {

    @Autowired
    private ExcelRepository excelRepository;

    public ExcelService(){

    }

    public ExcelService(String path) throws IOException {
        this.excelRepository = new ExcelRepository(path);
    }

    public int getNumberOfSheets() {
        return excelRepository.getWorkbook().getNumberOfSheets();
    }

    public List<Integer> getCaseNumberLenght(String sheetName, String caseNumber) {
        int sheetIndex = excelRepository.getSheetIndexBySheetName(sheetName);
        XSSFSheet sheet = excelRepository.getSheetBySheetIndex(sheetIndex);
        AtomicReference<Boolean> flag = new AtomicReference<>(false);
        List<Integer> caseNumStartEnd = new ArrayList<>();
        IntStream.rangeClosed(1, sheet.getLastRowNum())
                .forEach(rowNum -> {
                    if(sheet.getRow(rowNum).getCell(0).toString().equals(caseNumber) && !flag.get()){
                        caseNumStartEnd.add(rowNum);
                        flag.set(true);
                    } else if (!sheet.getRow(rowNum).getCell(0).toString().equals(caseNumber) && flag.get()){
                        caseNumStartEnd.add(rowNum-1);
                        flag.set(false);
                    }
                });
        return caseNumStartEnd;
    }

    public int getLastRowNumber(String sheetName){
        int sheetIndex = excelRepository.getSheetIndexBySheetName(sheetName);
        XSSFSheet sheet = excelRepository.getSheetBySheetIndex(sheetIndex);
        return sheet.getLastRowNum();
    }

    public Map<String, String> mapData(String sheetName, int headerIndex, int dataIndex) {
        int sheetIndex = excelRepository.getSheetIndexBySheetName(sheetName);
        XSSFSheet sheet = excelRepository.getSheetBySheetIndex(sheetIndex);

        List<String> headers = new ArrayList<>();
        List<String> data = new ArrayList<>();

        int rowLenght = sheet.getRow(headerIndex).getLastCellNum();

        IntStream.rangeClosed(1,rowLenght).forEach(round -> {
            headers.add(String.valueOf(sheet.getRow(headerIndex).getCell(round)));
            data.add(String.valueOf(sheet.getRow(dataIndex).getCell(round)));
        });

        Map<String, String> map = toMap(headers,data);
        return map;
    }

    public static  Map<String, String> toMap(List<String> keys, List<String> values) {
        return IntStream.range(0, keys.size()).boxed().collect(Collectors.toMap(keys::get, values::get));
    }
}
