Here's a README file for your project along with the necessary dependencies:

---

# Document Service API

This API allows you to upload Excel and Word files, replace placeholders in the Word document with values from the Excel file, and generate both Word and PDF documents.

## **API Endpoints**

### 1. **Upload Excel and Word Files**
**POST** `/api/v1/excel/upload`

This endpoint accepts two files, an Excel file (`excelFile`) and a Word template (`wordFile`), processes the Excel file, and generates new Word and PDF documents by replacing placeholders in the Word file with the data from the Excel file.

#### Request Parameters:
- **`excelFile`**: The Excel file (.xlsx) containing the data.
- **`wordFile`**: The Word template file (.docx) with placeholders to be replaced.

#### Response:
- A list of rows from the Excel file.

---

## **Dependencies**

To run this project, you need the following dependencies:

### **Spring Boot Dependencies**
Add the following dependencies to your `pom.xml` for the necessary libraries:

```xml
<dependencies>
    <!-- Spring Web for REST controllers -->
    <dependency>
        <groupId>org.springframework.boot</groupId>
        <artifactId>spring-boot-starter-web</artifactId>
    </dependency>
    
    <!-- Spring Boot for Dependency Injection -->
    <dependency>
        <groupId>org.springframework.boot</groupId>
        <artifactId>spring-boot-starter</artifactId>
    </dependency>

    <!-- Apache POI for Excel file parsing -->

    <dependency>
        <groupId>org.apache.poi</groupId>
        <artifactId>poi-ooxml</artifactId>
        <version>5.2.3</version>
    </dependency>



    <!-- Apache POI for Word file parsing -->


    <dependency>
        <groupId>com.itextpdf</groupId>
        <artifactId>itextpdf</artifactId>
        <version>5.5.13.3</version>
    </dependency>

    <!-- iText for PDF generation -->
    <dependency>
        <groupId>com.itextpdf</groupId>
        <artifactId>itextpdf</artifactId>
        <version>5.5.13.2</version>
    </dependency>

    <!-- Spring Boot Starter for File Uploads -->
    <dependency>
        <groupId>org.springframework.boot</groupId>
        <artifactId>spring-boot-starter-web</artifactId>
    </dependency>

    <!-- Spring Boot Starter for Thymeleaf (if needed for templates) -->
    <dependency>
        <groupId>org.springframework.boot</groupId>
        <artifactId>spring-boot-starter-thymeleaf</artifactId>
    </dependency>

    <!-- Optional: For Base64 encoding -->
    <dependency>
        <groupId>org.apache.commons</groupId>
        <artifactId>commons-codec</artifactId>
        <version>1.15</version>
    </dependency>
</dependencies>
```

### **Font for PDF Generation**
To properly render Persian text in the generated PDFs, make sure the font (`vazir.ttf`) is available in your resources folder under `src/main/resources/fonts/vazir.ttf`, if not push fonts folder in resources, in folder.


### **FPDF Generation Point**
First colume row consider as properties to change variables .
---

