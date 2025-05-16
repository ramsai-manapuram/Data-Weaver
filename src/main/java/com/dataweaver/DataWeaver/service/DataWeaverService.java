package com.dataweaver.DataWeaver.service;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;


@Service
public class DataWeaverService {

    public byte[] generateExcel(MultipartFile file) throws IOException {
        Workbook workbook = new XSSFWorkbook(file.getInputStream());
        Sheet sourceSheet = workbook.getSheetAt(0);

        // DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd-MMM-yyyy");
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("MMM dd, yyyy");
        Cell cell = sourceSheet.getRow(1).getCell(3);
        LocalDate date = LocalDate.parse(cell.toString(), formatter);

        int month = date.getMonth().getValue();
        int year = date.getYear();


        int rowCount = sourceSheet.getPhysicalNumberOfRows();
        

        for (int rowIndex = 1; rowIndex < rowCount; rowIndex++) {
            Cell dateCell = findDateFromSheet(sourceSheet, rowIndex);
            String inputDate = dateCell.toString();
            // Todo: convert from V2 to V1 format
            DateTimeFormatter inputFormatter = DateTimeFormatter.ofPattern("MMM dd, yyyy");
            LocalDate localDate = LocalDate.parse(inputDate, inputFormatter);

            DateTimeFormatter outputFormatter = DateTimeFormatter.ofPattern("dd-MMM-yyyy");
            String formattedDateStr = localDate.format(outputFormatter);
            dateCell.setCellValue(formattedDateStr);
        }

        TreeMap<String, Double> employeeNames = findAllEmployeeNames(sourceSheet);
        
        // for (Map.Entry<String, Double> entry: employeeNames.entrySet()) {
        //     System.out.println(entry.getKey() + " " + entry.getValue());
        // }

        Workbook outputWorkbook = new XSSFWorkbook();
        CellStyle borderStyle = outputWorkbook.createCellStyle();
        borderStyle.setBorderTop(BorderStyle.MEDIUM);
        borderStyle.setBorderBottom(BorderStyle.MEDIUM);
        borderStyle.setBorderLeft(BorderStyle.MEDIUM);
        borderStyle.setBorderRight(BorderStyle.MEDIUM);

        borderStyle.setAlignment(HorizontalAlignment.CENTER);
        borderStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        borderStyle.setWrapText(true);

        addSummaryPage(sourceSheet, outputWorkbook, employeeNames, borderStyle);
        addEachTimeSheet(outputWorkbook, employeeNames, sourceSheet, month, year, borderStyle);
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        outputWorkbook.write(outputStream);
        byte[] outputBytes = outputStream.toByteArray();

        return outputBytes;
    }

    private void addSummaryPage(Sheet sourceSheet, Workbook outputWorkbook, Map<String, Double> employeeNames, CellStyle borderStyle) {
        Sheet outputSheet = outputWorkbook.createSheet("Summary");
        String[] summaryColumns = {"Names", "Hours", "New/Existing"};
        addColumns(summaryColumns, outputSheet);
        fillSummarySheet(sourceSheet, outputSheet, employeeNames);
        fitColumnContent(summaryColumns.length, outputSheet);
        addBorders(outputSheet, borderStyle, summaryColumns.length);
        applyColour(outputWorkbook, outputSheet, summaryColumns.length, 0, IndexedColors.LIGHT_BLUE.getIndex());
    }

    private void addEachTimeSheet(Workbook outputWorkbook, Map<String, Double> employeeNames, Sheet sourceSheet, int month, int year, CellStyle style) {
        String[] columns = {"Name", "Date", "Title", "Description", "Project Time"};
        for (String name: employeeNames.keySet()) {
            Sheet currentSheet = outputWorkbook.createSheet(name);
            addColumns(columns, currentSheet);
            addEachPersonSheetData(outputWorkbook, style, sourceSheet, currentSheet, name, month, year, columns.length);
            fitColumnContent(columns.length, currentSheet);
            addBorders(currentSheet, style, columns.length);
            applyColour(outputWorkbook, currentSheet, columns.length, 0, IndexedColors.LIGHT_BLUE.getIndex());
            updateWeekendColour(outputWorkbook, currentSheet, columns.length);
            updateDateFormat(currentSheet);
        } 
    }

    private void updateDateFormat(Sheet sheet) {
        DateTimeFormatter outputFormatter = DateTimeFormatter.ofPattern("MMM dd, yyyy");
        // for (Row row: sheet) {
        //     Cell dateCell = row.getCell(1);

        //     LocalDate convertedDate = LocalDate.parse(dateCell.toString(), outputFormatter);
        //     dateCell.setCellValue(convertedDate.toString());

        // }
    }

    private void addBorders(Sheet sheet, CellStyle style, int totalColumns) {
        for (Row row: sheet) {
            for (int column = 0; column < totalColumns; column++) {
                if (row.getCell(column) != null) {
                    row.getCell(column).setCellStyle(style);
                }
            }
        }
    }

    private void updateWeekendColour(Workbook workbook, Sheet sheet, int length) {
        int rowIndex = 0;
        for (Row row: sheet) {
            if (rowIndex == 0) {
                rowIndex++;
                continue;
            }
            Cell cell = row.getCell(1);
            LocalDate date = LocalDate.parse(cell.toString());

            Cell descriptionCell = row.getCell(3);

            if (isWeekend(date)) {
                applyColour(workbook, sheet, length, rowIndex, IndexedColors.GREEN.getIndex());
            } else if (descriptionCell.toString().equals("")) {
                descriptionCell.setCellValue("On Leave");
                applyColour(workbook, sheet, length, rowIndex, IndexedColors.SKY_BLUE.getIndex());
            }
            rowIndex++;
        }
    }

    private void applyColour(Workbook outputWorkbook, Sheet sheet, int length, int rowIndex, short colourIndex) {
        CellStyle style = outputWorkbook.createCellStyle();
        style.setBorderTop(BorderStyle.MEDIUM);
        style.setBorderBottom(BorderStyle.MEDIUM);
        style.setBorderLeft(BorderStyle.MEDIUM);
        style.setBorderRight(BorderStyle.MEDIUM);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setWrapText(true);
        style.setFillForegroundColor(colourIndex);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        Row firstRow = sheet.getRow(rowIndex);
        for (int column = 0; column < length; column++) {
            Cell cell = firstRow.getCell(column);
            cell.setCellStyle(style);
        }
    }

    private void fitColumnContent(int length, Sheet sheet) {
        for (int column = 0; column < length; column++) {
            sheet.autoSizeColumn(column);
        }
    }

    private void addEachPersonSheetData(Workbook outputWorkbook, CellStyle style, Sheet sourceSheet, Sheet destinationSheet, String name, int month, int year, int length) {
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

    private int findColumnIndex(Sheet sourceSheet, String fieldName) {
        Row row = sourceSheet.getRow(0);
        int columnIndex = 0;
        for (Cell cell: row) {
            if (cell.toString().equals(fieldName)) {
                return columnIndex;
            }
            columnIndex++;
        }
        return -1;
    }

    private String findNameFromSheet(Sheet sourceSheet, int rowIndex) {
        int colIndex = findColumnIndex(sourceSheet, "Emp Name");
        return sourceSheet.getRow(rowIndex).getCell(colIndex).toString();
    }

    private Cell findDateFromSheet(Sheet sourceSheet, int rowIndex) {
        int colIndex = findColumnIndex(sourceSheet, "Date");
        return sourceSheet.getRow(rowIndex).getCell(colIndex);
    }

    private String findDescriptionFromSheet(Sheet sourceSheet, int rowIndex) {
        int colIndex = findColumnIndex(sourceSheet, "Description");
        return sourceSheet.getRow(rowIndex).getCell(colIndex).toString();
    }

    private void updateDescriptionAndHours(Sheet sourceSheet, Sheet destinationSheet, String name) {
        boolean isFirstRow = true;
        Set<String> allTasks = new HashSet<>();

        int rowCount = sourceSheet.getPhysicalNumberOfRows();

        for (int rowIndex = 1; rowIndex < rowCount; rowIndex++) {
            Row row = sourceSheet.getRow(rowIndex);
            if (row == null)    continue;

            int colCount = row.getPhysicalNumberOfCells();

            for (int colIndex = 0; colIndex < colCount; colIndex++) {
                Cell cell = row.getCell(colIndex);
                if (cell == null)   continue;

                String namePresent = findNameFromSheet(sourceSheet, rowIndex);
                if (!namePresent.equals(name)) {
                    continue;
                }

                Cell dateCell = findDateFromSheet(sourceSheet, rowIndex);
                int getDay = getDayFromDate(dateCell.toString());
                Cell descriptionCell = destinationSheet.getRow(getDay).getCell(3);
                Cell projectTimeCell = destinationSheet.getRow(getDay).getCell(4);

                String existingTask = descriptionCell.toString();
                String newTask = findDescriptionFromSheet(sourceSheet, rowIndex);
                if (allTasks.contains(newTask)) {
                    continue;
                }
                allTasks.add(newTask);
                if (existingTask.length() > 0) {
                    existingTask += ", ";
                }

                projectTimeCell.setCellValue("8");

                existingTask += newTask;
                descriptionCell.setCellValue(existingTask);
            }
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

        Cell secondCell = blankRow.createCell(1);
        Cell thirdCell = blankRow.createCell(2);

        Row totalHoursRow = destinationSheet.createRow(rowIndex++);
        Cell totalHoursFirstCol = totalHoursRow.createCell(0);
        totalHoursFirstCol.setCellValue("Total Hours");

        Cell totalHoursSecondCol = totalHoursRow.createCell(1);
        totalHoursSecondCol.setCellValue(Integer.toString(totalHours));

        Cell blankThirdCol = totalHoursRow.createCell(2);
    }

    private void addColumns(String[] columns, Sheet sheet) {
        Row row = sheet.createRow(0);
        int columnIndex = 0;
        for (String column: columns) {
            Cell cell = row.createCell(columnIndex++);
            cell.setCellValue(column);
        }
    }

    private int findEmployeeNameIndex(Sheet sheet) {
        Row firstRow = sheet.getRow(0);
        int columnIndex = 0;
        for (Cell column: firstRow) {
            if (column.toString().equals("Emp Name")) {
                return columnIndex;
            }
            columnIndex++;
        }

        return -1;
    }

    private int findTotalHoursIndex(Sheet sheet) {
        Row firstRow = sheet.getRow(0);
        int columnIndex = 0;
        for (Cell column: firstRow) {
            if (column.toString().equals("Total Hours")) {
                return columnIndex;
            }
            columnIndex++;
        }

        return -1;
    }

    private TreeMap<String, Double> findAllEmployeeNames(Sheet sheet) {
        TreeMap<String, Double> store = new TreeMap<>();
        Map<String, Set<Integer>> visited = new HashMap<>();
        int empNameIndex = findEmployeeNameIndex(sheet);
        int totalHoursIndex = findTotalHoursIndex(sheet);

        for (Row row: sheet) {
            Cell cell = row.getCell(empNameIndex);
            if (cell == null)   continue;
            Cell hoursCell = row.getCell(totalHoursIndex);
            String name = cell.toString();


            if (cell != null && !name.equals("Emp Name") && hoursCell != null) {

                Cell dateCell = row.getCell(3);
            
                int day = getDayFromDate(dateCell.toString());

                if (visited.containsKey(name) && visited.get(name).contains(day)) {
                    continue;
                }
                String hours = getCellValue(hoursCell);
                double hoursDouble = findHoursInDouble(hours);
                store.put(name, store.getOrDefault(name, 0.0) + hoursDouble);
                if (!visited.containsKey(name)) {
                    visited.put(name, new HashSet<>());
                }
                visited.get(name).add(day);
            }
        }

        return store;
    }

    private double findHoursInDouble(String hours) {
        int index = hours.indexOf(':');
        if (index != -1) {
            double result = Double.parseDouble(hours.substring(index - 2, index));
            int minutes = Integer.parseInt(hours.substring(index + 1, index + 3));
            if (minutes == 15) {
                result += 0.3;
            } else if (minutes == 30) {
                result += 0.5;
            } else if (minutes == 45) {
                result += 0.75;
            }
            return result;
        }

        return Double.parseDouble(hours);
    }

    private String getCellValue(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    return String.valueOf(cell.getNumericCellValue());
                }
            default:
                break;
        }
        return "";
    }

}
