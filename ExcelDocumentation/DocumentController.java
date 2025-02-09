package com.example.server.utill.ExcelDocumentation;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.util.List;
import java.util.Map;

@RestController
@RequestMapping("/api/v1/excel")
public class DocumentController {

    @Autowired
    private DocumentService documentService;

    public DocumentController(DocumentService documentService) {
        this.documentService = documentService;
    }

//    @PostMapping("/upload")
//    public ResponseEntity<List<Map<String, String>>> uploadExcel(@RequestParam("file") MultipartFile file) {
//        try {
//            List<Map<String, String>> rowsData = documentService.parseExcelFile(file);
//            return ResponseEntity.ok(rowsData);
//        } catch (RuntimeException e) {
//            return ResponseEntity.badRequest().build();
//        }
//    }


    @PostMapping("/upload")
    public ResponseEntity<List<Map<String, String>>> uploadExcel(
            @RequestParam("excelFile") MultipartFile excelFile,
            @RequestParam("wordFile") MultipartFile wordFile) {

        try {
            // Parse the Excel file
            List<Map<String, String>> rowsData = documentService.parseExcelFile(excelFile);

            // Iterate over the rows data and generate documents for each row
            for (int i = 0; i < rowsData.size(); i++) {
                Map<String, String> params = rowsData.get(i);
                documentService.generateDocument(params, wordFile);
            }

            return ResponseEntity.ok(rowsData);

        } catch (RuntimeException e) {
            e.printStackTrace();  // Log the exception for debugging purposes
            return ResponseEntity.badRequest().build();
        }
    }




    @PostMapping("/generate-doc")
    public ResponseEntity<Map<String, String>> generateDocument(
            @RequestParam Map<String, String> params,
            @RequestParam("wordFile") MultipartFile file) {

        try {
            Map<String, String> response = documentService.generateDocument(params, file);
            return ResponseEntity.ok(response);
        } catch (RuntimeException e) {
            return ResponseEntity.internalServerError().build();
        }
    }


    


}
