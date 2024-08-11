package com.excel.controller;


import com.excel.exception.EmptyExcelFileException;
import com.excel.exception.InvalidExcelFileException;
import jakarta.validation.Valid;
import jakarta.validation.constraints.NotNull;
import java.util.regex.Pattern;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.InputStreamResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;

@RestController
@RequestMapping("/api/excel")
public class ExcelController {

  private static final String XLSX_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  private static final String XLS_CONTENT_TYPE = "application/vnd.ms-excel";
  private static final Pattern validPattern = Pattern.compile("^[a-zA-Z0-9\\s]*$");


  @PostMapping("/upload")
  public ResponseEntity<?> uploadExcel(@RequestParam("file") MultipartFile file) throws IOException {

    if (file == null || file.isEmpty()) {
      return ResponseEntity.badRequest().body("Uploaded Excel file is empty.");
    }
    // Validate file type
    String contentType = file.getContentType();
    if (!XLSX_CONTENT_TYPE.equals(contentType) && !XLS_CONTENT_TYPE.equals(contentType)) {
      throw new InvalidExcelFileException("Uploaded file is not a excel file");
    }
    // Process the file
    var cleanedExcelOutputStream = new ByteArrayOutputStream();
    processExcelFile(file, cleanedExcelOutputStream);

    // Prepare response with cleaned Excel and statistics
    var inputStream = new ByteArrayInputStream(cleanedExcelOutputStream.toByteArray());
    Resource resource = new InputStreamResource(inputStream);

    var headers = new HttpHeaders();
    headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
    headers.set(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=cleaned_file.xlsx");

    return new ResponseEntity<>(resource, headers, HttpStatus.OK);
  }

  private static void processExcelFile(MultipartFile file, ByteArrayOutputStream cleanedExcelOutputStream) throws IOException {

    Workbook workbook = new XSSFWorkbook(file.getInputStream());
    if (workbook.getNumberOfSheets() == 0) {
      throw new EmptyExcelFileException("The Excel file is empty.");
    }
    var sheet = workbook.getSheetAt(0);

    var totalCells = 0;
    var emptyCells = 0;
    var corruptedCells = 0;
    var validCells = 0;

    for (Row row : sheet) {
      int lastCellNum = row.getLastCellNum();
      for (var cellNum = 0; cellNum < lastCellNum; cellNum++) {
        var cell = row.getCell(cellNum);
        totalCells++;
        if (cell == null || cell.getCellType() == CellType.BLANK) {
          emptyCells++;
          shiftCellsLeft(row, cellNum, lastCellNum);
          // Decrease lastCellNum as cells are shifted
          lastCellNum--;
          // Re-check the current cell after shifting
          cellNum--;
        } else if (cell.getCellType() == CellType.STRING) {
          var cellValue = cell.getStringCellValue();
          if (!validPattern.matcher(cellValue).matches()) {
            corruptedCells++;
            shiftCellsLeft(row, cellNum, lastCellNum);
            lastCellNum--;
            cellNum--;
          } else {
            validCells++;
          }
        } else if (cell.getCellType() == CellType.NUMERIC) {
          validCells++;
        } else {
          corruptedCells++;
          shiftCellsLeft(row, cellNum, lastCellNum);
          lastCellNum--;
          cellNum--;
        }
      }
    }

    // Add statistics to a new sheet
    var statsSheet = workbook.createSheet("Statistics");
    var statsRow = statsSheet.createRow(0);
    statsRow.createCell(0).setCellValue("Total Cells Processed");
    statsRow.createCell(1).setCellValue(totalCells);
    statsRow = statsSheet.createRow(1);
    statsRow.createCell(0).setCellValue("Valid Cells");
    statsRow.createCell(1).setCellValue(validCells);
    statsRow = statsSheet.createRow(2);
    statsRow.createCell(0).setCellValue("Empty Cells");
    statsRow.createCell(1).setCellValue(emptyCells);
    statsRow = statsSheet.createRow(3);
    statsRow.createCell(0).setCellValue("Corrupted Cells");
    statsRow.createCell(1).setCellValue(corruptedCells);

    // Write the cleaned workbook to the output stream
    workbook.write(cleanedExcelOutputStream);
    workbook.close();

  }

  private static void shiftCellsLeft(Row row, int start, int last) {
    for (int i = start; i < last - 1; i++) {
      var currentCell = row.getCell(i);
      var nextCell = row.getCell(i + 1);

      if (nextCell != null) {
        if (currentCell == null) {
          currentCell = row.createCell(i, nextCell.getCellType());
        }
        cloneCell(currentCell, nextCell);
      } else if (currentCell != null) {
        row.removeCell(currentCell);
      }
    }

    // Remove the last cell after shifting
    var lastCell = row.getCell(last - 1);
    if (lastCell != null) {
      row.removeCell(lastCell);
    }
  }

  private static void cloneCell(Cell newCell, Cell oldCell) {
    switch (oldCell.getCellType()) {
      case STRING:
        newCell.setCellValue(oldCell.getStringCellValue());
        break;
      case NUMERIC:
        newCell.setCellValue(oldCell.getNumericCellValue());
        break;
      case BOOLEAN:
        newCell.setCellValue(oldCell.getBooleanCellValue());
        break;
      case FORMULA:
        newCell.setCellFormula(oldCell.getCellFormula());
        break;
      default:
        break;
    }
  }

}

