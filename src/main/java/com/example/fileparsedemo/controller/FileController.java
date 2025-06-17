package com.example.fileparsedemo.controller;

import com.example.fileparsedemo.model.ResultResponse;
import com.example.fileparsedemo.service.FileFullService;
import com.example.fileparsedemo.service.FileLimitService;

import lombok.AllArgsConstructor;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

@RestController
@AllArgsConstructor
public class FileController {
    private final FileFullService fileFullService;
    private final FileLimitService fileLimitService;

    @PostMapping("/api/v1/file/download")
    public ResponseEntity<?> importFile(@RequestParam("file") MultipartFile file){
        try {
            byte[] excelData;
            long sizeInBytes = file.getSize();
            long maxLimit = 2 * 1024 * 1024; // 2MB
            if(sizeInBytes <= maxLimit){
                ResultResponse result = fileLimitService.readExcelFile(file);
                excelData = fileLimitService.writeExcelFile(file, result);
            } else{
                ResultResponse result = fileFullService.readFileXlsx(file);
                excelData = fileFullService.writeFileXlsx(result);
            }
            HttpHeaders headers = new HttpHeaders();
            headers.setContentType(MediaType.parseMediaType(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ));
            headers.setContentDispositionFormData("attachment", "ket_qua.xlsx");
            return new ResponseEntity<>(excelData, headers, HttpStatus.OK);
        } catch (Exception e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).build();
        }
    }

    @PostMapping("/api/v2/file/download")
    public ResponseEntity<byte[]> downloadExcel(@RequestParam("file") MultipartFile file) {
        try {
            ResultResponse result = fileFullService.readFileXlsx(file);
            byte[] excelData = fileFullService.writeFileXlsx(result);
            HttpHeaders headers = new HttpHeaders();
            headers.setContentType(MediaType.parseMediaType(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ));
            headers.setContentDispositionFormData("attachment", "ket_qua.xlsx");
            return new ResponseEntity<>(excelData, headers, HttpStatus.OK);
        } catch (Exception e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).build();
        }
    }

    @GetMapping("/test")
    public String test(){
        return "tesssssss";
    }
}
