package com.project.excel_diff.controller;

import com.project.excel_diff.service.ExcelToolsService;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;

import lombok.AllArgsConstructor;


@RestController
@RequestMapping("/api/excelTools")
@AllArgsConstructor
public class ExcelToolController {

    private final ExcelToolsService excelToolsService;


    @PostMapping("/differences")
    public ResponseEntity<InputStreamResource> getExcelSheetsDifferences(@RequestParam("file") MultipartFile file) throws Exception {
        try {
            ByteArrayOutputStream byteArrayOutputStream = excelToolsService.findDifferencesBetweenSheets(file);

            ByteArrayInputStream byteArrayInputStream = new ByteArrayInputStream(byteArrayOutputStream.toByteArray());
            HttpHeaders headers = new HttpHeaders();
            headers.add("Content-Disposition", "attachment; filename=output_diff.xlsx");

            return ResponseEntity.ok()
                    .headers(headers)
                    .contentType(MediaType.APPLICATION_OCTET_STREAM)
                    .body(new InputStreamResource(byteArrayInputStream));
        } catch (Exception e) {
            throw new Exception("Exception from getExcelSheetsDifferences " + e.getMessage());
        }
    }

}
