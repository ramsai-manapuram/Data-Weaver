package com.dataweaver.DataWeaver.service;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import com.dataweaver.DataWeaver.exception.CustomException;

@Service
public class DataWeaverService {

    public byte[] generateExcel(MultipartFile file, int month, int year) throws IOException {
        if (month < 1 || month > 12) {
            throw new CustomException("Invalid month passed");
        } 
        if (year < 2000 || year > 2050) {
            throw new CustomException("Invalid year passed");
        }
        Workbook workbook = new XSSFWorkbook(file.getInputStream());
        Sheet sourceSheet = workbook.getSheetAt(0);
        Map<String, Double> employeeNames = findAllEmployeeNames(sourceSheet);

        Workbook outputWorkbook = new XSSFWorkbook();
        Sheet outputSheet = outputWorkbook.createSheet("Summary");

        addSummaryPage(sourceSheet, outputSheet, employeeNames);
        addEachTimeSheet(outputWorkbook, employeeNames, sourceSheet, month, year);

        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        outputWorkbook.write(outputStream);
        byte[] outputBytes = outputStream.toByteArray();

        return outputBytes;
    }

    private void addSummaryPage(Sheet sourceSheet, Sheet outputSheet, Map<String, Double> employeeNames) {
        String[] summaryColumns = {"Names", "Hours", "New/Existing"};
        addColumns(summaryColumns, outputSheet);
        fillSummarySheet(sourceSheet, outputSheet, employeeNames);
        fitColumnContent(summaryColumns.length, outputSheet);
    }

    private void addEachTimeSheet(Workbook outputWorkbook, Map<String, Double> employeeNames, Sheet sourceSheet, int month, int year) {
        String[] columns = {"Name", "Date", "Title", "Description", "Project Time"};

        for (String name: employeeNames.keySet()) {
            Sheet currentSheet = outputWorkbook.createSheet(name);
            addColumns(columns, currentSheet);
            addEachPersonSheetData(sourceSheet, currentSheet, name, month, year);
            fitColumnContent(columns.length, currentSheet);
        }
    }

    private void fitColumnContent(int length, Sheet sheet) {
        for (int column = 0; column < length; column++) {
            sheet.autoSizeColumn(column);
        }
    }

    private void addEachPersonSheetData(Sheet sourceSheet, Sheet destinationSheet, String name, int month, int year) {
        LocalDate firstDate = LocalDate.of(year, month, 1);
        LocalDate lastDate = firstDate.withDayOfMonth(firstDate.lengthOfMonth());

        int rowIndex = 1;
        for (LocalDate date = firstDate; !date.isAfter(lastDate); date = date.plusDays(1)) {
            Row row = destinationSheet.createRow(rowIndex++);
            Cell nameCell = row.createCell(0);
            nameCell.setCellValue(name);
            Cell dateCell = row.createCell(1);
            dateCell.setCellValue(date.toString());
            Cell titleCell = row.createCell(2);

            Cell descriptionCell = row.createCell(3);
            Cell projectTimeCell = row.createCell(4);
            
            if (!isWeekend(date)) {
                titleCell.setCellValue("Development");
            }
        }

        updateDescriptionAndHours(sourceSheet, destinationSheet, name);
    }

    private void updateDescriptionAndHours(Sheet sourceSheet, Sheet destinationSheet, String name) {
        boolean isFirstRow = true;
        Set<String> allTasks = new HashSet<>();

        for (Row row: sourceSheet) {
            if (isFirstRow) {
                isFirstRow = false;
                continue;
            }
            
            if (!name.equals(row.getCell(1).toString())) {
                continue;
            }
            Cell dateCell = row.getCell(3);
            
            int getDay = getDayFromDate(dateCell.toString());
            Cell descriptionCell = destinationSheet.getRow(getDay).getCell(3);
            String existingTask = descriptionCell.toString();

            String newTask = row.getCell(5).toString();
            if (allTasks.contains(newTask)) {
                continue;
            }
            allTasks.add(newTask);
            if (existingTask.length() > 0) {
                existingTask += ", ";
            }
           
            existingTask += newTask;
            descriptionCell.setCellValue(existingTask);
        }
    }

    private int getDayFromDate(String date) {
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd-MMM-yyyy");
        try {
            LocalDate parsedDate = LocalDate.parse(date, formatter);
            return parsedDate.getDayOfMonth();
        } catch (DateTimeParseException e) {
            System.out.println("Invalid date format: " + date);
            throw e;
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
