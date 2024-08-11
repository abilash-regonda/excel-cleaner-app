package com.excel;

import java.util.Objects;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.boot.test.web.client.TestRestTemplate;
import org.springframework.boot.test.web.server.LocalServerPort;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpEntity;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpMethod;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.util.LinkedMultiValueMap;
import org.springframework.util.MultiValueMap;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;

import static org.assertj.core.api.Assertions.assertThat;

@SpringBootTest(webEnvironment = SpringBootTest.WebEnvironment.RANDOM_PORT)
class ExcelControllerTest {

  @LocalServerPort
  private int port;

  @Autowired
  private TestRestTemplate restTemplate;

  @Test
  void testProcessExcelFile() throws IOException {
    // Create a test Excel file
    Workbook workbook = new XSSFWorkbook();
    Sheet sheet = workbook.createSheet("TestSheet");

    Row row1 = sheet.createRow(0);
    row1.createCell(0).setCellValue("Valid1");
    row1.createCell(1).setCellValue("oemõŸÿüwÜq‡TÿLÝ¢¬ß¿‡I£†");
    row1.createCell(3).setCellValue("123abc");

    Row row2 = sheet.createRow(1);
    row2.createCell(0).setCellValue("1234");
    row2.createCell(2).setCellValue("Invalid\\xFF\\xFE");
    row2.createCell(3).setCellValue("Valid4");

    ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
    workbook.write(outputStream);
    workbook.close();
    byte[] excelBytes = outputStream.toByteArray();

    // Create a MultiValueMap to hold the file to be uploaded
    MultiValueMap<String, Object> body = new LinkedMultiValueMap<>();
    HttpHeaders headers = new HttpHeaders();
    headers.setContentType(MediaType.MULTIPART_FORM_DATA);

    // Add the Excel file as a ByteArrayResource
    body.add("file", new ByteArrayResource(excelBytes) {
      @Override
      public String getFilename() {
        return "test.xlsx";
      }
    });

    HttpEntity<MultiValueMap<String, Object>> requestEntity = new HttpEntity<>(body, headers);

    // Send the POST request to the API
    ResponseEntity<Resource> responseEntity = restTemplate.exchange(
        "http://localhost:" + port + "/api/excel/upload",
        HttpMethod.POST,
        requestEntity,
        Resource.class
    );

    // Validate the response
    assertThat(responseEntity.getStatusCode()).isEqualTo(HttpStatus.OK);
    assertThat(responseEntity.getHeaders().getContentType()).isEqualTo(MediaType.APPLICATION_OCTET_STREAM);

    // Additional validation on the returned Excel file
    try (Workbook cleanedWorkbook = new XSSFWorkbook(new ByteArrayInputStream(
        Objects.requireNonNull(responseEntity.getBody()).getInputStream().readAllBytes()))) {
      Sheet cleanedSheet = cleanedWorkbook.getSheetAt(0);

      // Check that the cells were shifted left correctly
      Row cleanedRow1 = cleanedSheet.getRow(0);
      assertThat(cleanedRow1.getCell(0).getStringCellValue()).isEqualTo("Valid1");
      assertThat(cleanedRow1.getCell(1).getStringCellValue()).isEqualTo("123abc");
      assertThat(cleanedRow1.getCell(2)).isNull(); // The cell should be null

      Row cleanedRow2 = cleanedSheet.getRow(1);
      assertThat(cleanedRow2.getCell(0).getStringCellValue()).isEqualTo("1234");
      assertThat(cleanedRow2.getCell(1).getStringCellValue()).isEqualTo("Valid4");
      assertThat(cleanedRow2.getCell(2)).isNull(); // The cell should be null

      // Validate the statistics in the second sheet
      Sheet statsSheet = cleanedWorkbook.getSheet("Statistics");
      assertThat(statsSheet).isNotNull();

      Row totalRow = statsSheet.getRow(0); // Assuming row 1 contains the values
      int totalCells = (int) totalRow.getCell(1).getNumericCellValue();
      Row validRow = statsSheet.getRow(1);
      int validCells = (int) validRow.getCell(1).getNumericCellValue();
      Row emptyRow = statsSheet.getRow(2);
      int emptyCells = (int) emptyRow.getCell(1).getNumericCellValue();
      Row corruptedRow = statsSheet.getRow(3);
      int corruptedCells = (int) corruptedRow.getCell(1).getNumericCellValue();

      // Assert the statistics match expected values
      assertThat(totalCells).isEqualTo(8);
      assertThat(validCells).isEqualTo(4);
      assertThat(emptyCells).isEqualTo(2);
      assertThat(corruptedCells).isEqualTo(2);
    }
  }

  @Test
  void testProcessInvalidExcelFile() {
    // Create a corrupted Excel file (e.g., missing data or invalid format)
    byte[] corruptedExcelBytes = new byte[]{0x50, 0x4B, 0x03, 0x04}; // Random bytes to simulate corruption

    // Create a MultiValueMap to hold the file to be uploaded
    MultiValueMap<String, Object> body = new LinkedMultiValueMap<>();
    HttpHeaders headers = new HttpHeaders();
    headers.setContentType(MediaType.MULTIPART_FORM_DATA);

    // Add the corrupted Excel file as a ByteArrayResource
    body.add("file", new ByteArrayResource(corruptedExcelBytes) {
      @Override
      public String getFilename() {
        return "test.pdf";
      }
    });

    HttpEntity<MultiValueMap<String, Object>> requestEntity = new HttpEntity<>(body, headers);

    // Send the POST request to the API
    ResponseEntity<String> responseEntity = restTemplate.exchange(
        "http://localhost:" + port + "/api/excel/upload",
        HttpMethod.POST,
        requestEntity,
        String.class
    );

    // Validate the response
    assertThat(responseEntity.getStatusCode()).isEqualTo(HttpStatus.BAD_REQUEST);
  }

  @Test
  void testProcessEmptyExcelFile() throws IOException {

    Workbook workbook = new XSSFWorkbook();
    ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
    workbook.write(outputStream);
    workbook.close();
    byte[] emptyExcelBytes = outputStream.toByteArray();

    // Create a MultiValueMap to hold the file to be uploaded
    MultiValueMap<String, Object> body = new LinkedMultiValueMap<>();
    HttpHeaders headers = new HttpHeaders();
    headers.setContentType(MediaType.MULTIPART_FORM_DATA);

    // Add the corrupted Excel file as a ByteArrayResource
    body.add("file", new ByteArrayResource(emptyExcelBytes) {
      @Override
      public String getFilename() {
        return "empty.xlsx";
      }
    });

    HttpEntity<MultiValueMap<String, Object>> requestEntity = new HttpEntity<>(body, headers);

    // Send the POST request to the API
    ResponseEntity<String> responseEntity = restTemplate.exchange(
        "http://localhost:" + port + "/api/excel/upload",
        HttpMethod.POST,
        requestEntity,
        String.class
    );

    // Validate the response
    assertThat(responseEntity.getStatusCode()).isEqualTo(HttpStatus.BAD_REQUEST);
  }


}
