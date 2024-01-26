package com.example;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader {

  public List<Customer> readExcelFile(String filePath) {
    List<Customer> dataList = new ArrayList<>();

    try (FileInputStream fis = new FileInputStream(new File(filePath));
        XSSFWorkbook workbook = new XSSFWorkbook(fis)) {

      Sheet sheet = workbook.getSheetAt(0);

      for (Row row : sheet) {
        if (row.getRowNum() == 0) { // Skippar första raden (rubrikerna)
          continue;
        }

        // Läs data från varje cell i raden och skapa ett Customer-objekt
        Integer customerID = safeGetIntCellValue(row, 0);
        String firstName = safeGetStringCellValue(row, 1);
        String lastName = safeGetStringCellValue(row, 2);
        String phoneNumber = safeGetStringCellValue(row, 3);
        String address = safeGetStringCellValue(row, 4);
        String payingCustomer = safeGetStringCellValue(row, 5);
        String doNotContact = safeGetStringCellValue(row, 6);

        Customer customer = new Customer(customerID, firstName, lastName,
            phoneNumber, address, payingCustomer,
            doNotContact);
        dataList.add(customer);
      }

    } catch (Exception e) {
      e.printStackTrace();
    }

    return dataList;
  }

  private Integer safeGetIntCellValue(Row row, int cellIndex) {
    Cell cell = row.getCell(cellIndex);
    return (cell != null) ? (int) cell.getNumericCellValue() : null;
  }

  private String safeGetStringCellValue(Row row, int cellIndex) {
    Cell cell = row.getCell(cellIndex);
    return (cell != null) ? cell.getStringCellValue() : "";
  }
}