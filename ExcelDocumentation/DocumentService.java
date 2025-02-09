package com.example.server.utill.ExcelDocumentation;


import com.itextpdf.text.Document;
import com.itextpdf.text.Element;
import com.itextpdf.text.Font;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfWriter;
import org.springframework.web.multipart.MultipartFile;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.springframework.stereotype.Service;

import java.io.*;
import java.util.*;
import java.util.Base64;

@Service
public class DocumentService {


    private String uploadDir="src/main/resources/static/";

    public List<Map<String, String>> parseExcelFile(MultipartFile file) {
        List<Map<String, String>> rowsData = new ArrayList<>();
        List<String> headers = new ArrayList<>();

        try (XSSFWorkbook workbook = new XSSFWorkbook(file.getInputStream())) {
            Sheet sheet = workbook.getSheetAt(0); // Get the first sheet

            // Read header row (first row)
            Row headerRow = sheet.getRow(0);
            for (Cell cell : headerRow) {
                headers.add(getCellValue(cell)); // Add header names
            }

            // Read data rows
            for (int i = 1; i <= sheet.getLastRowNum(); i++) { // Start from the second row
                Row row = sheet.getRow(i);
                Map<String, String> rowData = new HashMap<>();

                for (String header : headers) {
                    int columnIndex = headers.indexOf(header);
                    Cell cell = row.getCell(columnIndex);
                    String cellValue = (cell != null) ? getCellValue(cell) : "تعریف نشده"; // Check if cell is null
                    rowData.put(header, cellValue); // Use header as key
                }

                rowsData.add(rowData);
            }
        } catch (IOException e) {
            throw new RuntimeException("Error parsing Excel file", e);
        }

        return rowsData;
    }

    private String getCellValue(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }


    // Seperate file

    public Map<String, String> generateDocument(Map<String, String> params, MultipartFile file) {
        try {
            // بارگذاری سند Word از فایل ورودی
            InputStream templateStream = file.getInputStream();
            XWPFDocument document = new XWPFDocument(templateStream);

            // جایگزینی مقادیر در سند Word
            for (XWPFParagraph paragraph : document.getParagraphs()) {
                for (XWPFRun run : paragraph.getRuns()) {
                    String text = run.getText(0);
                    if (text != null) {
                        for (Map.Entry<String, String> entry : params.entrySet()) {
                            text = text.replace("{{" + entry.getKey() + "}}", entry.getValue());
                        }
                        run.setText(text, 0);
                    }
                }
            }

            // تولید نام مشترک برای فایل‌ها
            String baseFileName = "generated_document_" + System.currentTimeMillis();

            // ذخیره فایل Word
            String wordFileName = baseFileName + ".docx";
            File outputWordFile = new File(uploadDir + "/" + wordFileName);
            try (FileOutputStream fileOut = new FileOutputStream(outputWordFile)) {
                document.write(fileOut);
            }

            // تبدیل فایل Word به PDF
            String pdfFileName = baseFileName + ".pdf";
            File outputPdfFile = new File(uploadDir + "/" + pdfFileName);

            // استفاده از iText برای ایجاد فایل PDF
            Document pdfDocument = new Document();
            PdfWriter.getInstance(pdfDocument, new FileOutputStream(outputPdfFile));
            pdfDocument.open();

            // افزودن فونت فارسی برای پشتیبانی از زبان فارسی در iText
            String fontPath = "src/main/resources/fonts/vazir.ttf"; // مسیر فونت فارسی
            BaseFont baseFont = BaseFont.createFont(fontPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            Font farsiFont = new Font(baseFont, 12, Font.NORMAL);

            // افزودن محتوای پاراگراف‌ها به PDF
            for (XWPFParagraph paragraph : document.getParagraphs()) {
                String paragraphText = paragraph.getText();
                if (paragraphText != null && !paragraphText.trim().isEmpty()) {
                    Paragraph pdfParagraph = new Paragraph(paragraphText, farsiFont);
                    pdfParagraph.setAlignment(Element.ALIGN_RIGHT); // راست‌چین کردن متن
                    pdfDocument.add(pdfParagraph);
                }
            }
            pdfDocument.close();

            // تبدیل سند Word به آرایه بایت و رمزگذاری به Base64
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            document.write(out);
            document.close();
            byte[] documentContent = out.toByteArray();
            String base64Encoded = Base64.getEncoder().encodeToString(documentContent);

            // آماده‌سازی پاسخ
            Map<String, String> response = new HashMap<>();
            response.put("base64Word", base64Encoded);
            response.put("fileUrlWord", "/static/" + wordFileName);
            response.put("fileUrlPdf", "/static/" + pdfFileName);

            return response;
        } catch (Exception e) {
            throw new RuntimeException("Error generating document", e);
        }
    }



}

