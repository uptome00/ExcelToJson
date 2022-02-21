package com.example.ExcelToJson.controller;

import com.example.ExcelToJson.service.ExcelService;
import lombok.Data;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;

import java.io.IOException;
import java.net.URI;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.IntStream;


@Controller
@Slf4j
@Data
public class ExcelController {

    @Autowired
    private ExcelService excelService;


    public HttpResponse<String> sendApi(Map data) throws IOException, OpenXML4JException, InterruptedException {
        HttpClient client = HttpClient.newHttpClient();
        HttpRequest request = HttpRequest.newBuilder()
                .uri(URI.create("http://localhost:4001/user"))
                .GET()
                .build();
        HttpResponse<String> response = client.send(request, HttpResponse.BodyHandlers.ofString());
        return response;
    }


    public void allCase(String path, String sheetName) throws  IOException  {
        excelService = new ExcelService(path);
        int lastRow = excelService.getLastRowNumber(sheetName);
        IntStream.rangeClosed(1,lastRow).forEach(round -> {
            Map<String,String> map = excelService.mapData(sheetName,0 ,round);
            try {
                sendApi(map);
            } catch (IOException e) {
                e.printStackTrace();
            } catch (OpenXML4JException e) {
                e.printStackTrace();
            } catch (InterruptedException e) {
                e.printStackTrace();
            }
        });


    }

    public void selectCase(String path, String sheetName, String caseNumber) throws OpenXML4JException, IOException, InterruptedException {
        excelService = new ExcelService(path);
        List<Integer> a  = excelService.getCaseNumberLenght(sheetName, caseNumber);
        log.info(String.valueOf(a));
        IntStream.rangeClosed(a.get(0),a.get(1)).forEach(round -> {
            Map<String,String> map = excelService.mapData(sheetName,0 ,round);
            try {
                sendApi(map);
            } catch (IOException e) {
                e.printStackTrace();
            } catch (OpenXML4JException e) {
                e.printStackTrace();
            } catch (InterruptedException e) {
                e.printStackTrace();
            }
        });
    }


}

@Data
class Path{
    private String path;
}
