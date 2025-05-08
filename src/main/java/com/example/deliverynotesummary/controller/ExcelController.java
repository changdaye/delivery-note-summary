package com.example.deliverynotesummary.controller;

import com.example.deliverynotesummary.service.ExcelService;
import lombok.RequiredArgsConstructor;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.List;

@RestController
@RequestMapping("/api/excel")
@RequiredArgsConstructor
public class ExcelController {

    private final ExcelService excelService;

    @PostMapping("/merge")
    public ResponseEntity<byte[]> mergeExcelFiles(@RequestParam("files") List<MultipartFile> files) throws IOException {
        byte[] mergedContent = excelService.mergeExcelFiles(files);
        
        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
        headers.setContentDispositionFormData("attachment", "merged_excel.xlsx");
        
        return ResponseEntity.ok()
                .headers(headers)
                .body(mergedContent);
    }
}