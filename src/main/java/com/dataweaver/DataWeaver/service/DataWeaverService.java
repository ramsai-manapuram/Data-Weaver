package com.dataweaver.DataWeaver.service;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

@Service
public class DataWeaverService {

    public byte[] generateExcel(MultipartFile file) throws IOException {
        Workbook workbook = new XSSFWorkbook(file.getInputStream());
        Sheet sourceSheet = workbook.getSheetAt(0);
        Map<String, Double> employeeNames = findAllEmployeeNames(sourceSheet);

        Workbook outputWorkbook = new XSSFWorkbook();
        Sheet outputSheet = outputWorkbook.createSheet("Summary");
        String[] summaryColumns = {"Names", "Hours", "New/Existing"};
        addColumns(summaryColumns, outputSheet);
        fillSummarySheet(sourceSheet, outputSheet, employeeNames);

        String[] columns = {"Name", "Date", "Title", "Description", "Project Time"};

        for (String name: employeeNames.keySet()) {
            Sheet currentSheet = outputWorkbook.createSheet(name);
            addColumns(columns, currentSheet);
            addEachPersonSheetData(sourceSheet, currentSheet, name);
        }

        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        outputWorkbook.write(outputStream);
        byte[] outputBytes = outputStream.toByteArray();

        return outputBytes;
    }

    private void addEachPersonSheetData(Sheet sourceSheet, Sheet destinationSheet, String name) {
        LocalDate firstDate = LocalDate.of(2025, 3, 1);
        LocalDate lastDate = firstDate.withDayOfMonth(firstDate.lengthOfMonth());

        int rowIndex = 1;
        for (LocalDate date = firstDate; !date.isAfter(lastDate); date = date.plusDays(1)) {
            Row row = destinationSheet.createRow(rowIndex++);
            Cell nameCell = row.createCell(0);
            nameCell.setCellValue(name);
            Cell dateCell = row.createCell(1);
            dateCell.setCellValue(date.toString());
            Cell titleCell = row.createCell(2);
            
            if (!isWeekend(date)) {
                titleCell.setCellValue("Development");
            }
        }
    }

    private boolean isWeekend(LocalDate date) {
        DayOfWeek dayOfWeek = date.getDayOfWeek();
        return dayOfWeek == DayOfWeek.SATURDAY || dayOfWeek == DayOfWeek.SUNDAY;
    }

    private void fillSummarySheet(Sheet sourceSheet, Sheet destinationSheet, Map<String, Double> employeeNames) {
        int rowIndex = 1;
        int totalHours = 0;

        for (String name: employeeNames.keySet()) {
            Row row = destinationSheet.createRow(rowIndex++);
            Cell cell = row.createCell(0);
            cell.setCellValue(name);

            double hours = employeeNames.get(name);
            totalHours += hours;
            Cell hoursCell = row.createCell(1);
            hoursCell.setCellValue(Double.toString(hours));

            Cell thirdCol = row.createCell(2);
            thirdCol.setCellValue("Existing");
        }

        Row blankRow = destinationSheet.createRow(rowIndex++);
        Cell blankCell = blankRow.createCell(0);

        Row totalHoursRow = destinationSheet.createRow(rowIndex++);
        Cell totalHoursFirstCol = totalHoursRow.createCell(0);
        totalHoursFirstCol.setCellValue("Total Hours");

        Cell totalHoursSecondCol = totalHoursRow.createCell(1);
        totalHoursSecondCol.setCellValue(Integer.toString(totalHours));
    }

    private void addColumns(String[] columns, Sheet sheet) {
        Row row = sheet.createRow(0);
        int columnIndex = 0;
        for (String column: columns) {
            Cell cell = row.createCell(columnIndex++);
            cell.setCellValue(column);
        }
    }

    private String getCellData(Cell cell) {
        String cellValue = "";
        switch (cell.getCellType()) {
            case STRING:
                cellValue = cell.getStringCellValue();
                break;

            case NUMERIC:
                cellValue = String.valueOf(cell.getNumericCellValue());
                break;

            case BOOLEAN:
                cellValue = String.valueOf(cell.getBooleanCellValue());
                break;

            case BLANK:
                cellValue = "";
                break;

            case ERROR:
                cellValue = "Error";
                break;

            default:
                cellValue = "Unknown";
        }
        return cellValue;
    }

    private Map<String, Double> findAllEmployeeNames(Sheet sheet) {
        Map<String, Double> store = new HashMap<>();

        for (Row row: sheet) {
            Cell cell = row.getCell(1);
            Cell hoursCell = row.getCell(6);
            String name = cell.toString();

            if (cell != null && !name.equals("Emp Name")) {
                double hours = Double.parseDouble(hoursCell.toString());
                store.put(name, store.getOrDefault(name, 0.0) + hours);
            }
        }

        return store;
    }

}
