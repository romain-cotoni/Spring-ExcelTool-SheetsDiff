package com.project.excel_diff.service;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.springframework.web.multipart.MultipartFile;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

@Service
public class ExcelToolsService1 {


    public ByteArrayOutputStream findDifferencesBetweenSheets(MultipartFile file) throws Exception {

        try {
            Workbook workbook = new XSSFWorkbook(file.getInputStream());

            List<Map<String, List<String>>> excelSheetsData = this.extractSheetsDataFromFile(file);

            StringBuilder summaryToDisplay = new StringBuilder();

            // Parse all sheets
            for (int sheetIndex = 1; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) { /*sheetIndex < excelSheets.size()*/
                String differencesToDisplay = compareRows(excelSheetsData.get(sheetIndex - 1),
                                                          excelSheetsData.get(sheetIndex),
                                                          workbook.getSheetAt(sheetIndex - 1),
                                                          workbook.getSheetAt(sheetIndex));
                summaryToDisplay.append("Différences entre tab ").append(sheetIndex).append(" et tab ").append(sheetIndex + 1).append(":\n")
                        .append(differencesToDisplay)
                        .append("\n\n");
            }

            // Display differences as text
            System.out.println(summaryToDisplay); // For future functionality to display text

            // Save into byteArrayOutputStream an updated Excel file with highlighted and labeled differences
            ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
            workbook.write(byteArrayOutputStream);
            workbook.close();

            return byteArrayOutputStream;
        } catch (Exception e) {
            throw new Exception("Exception from findDifferencesBetweenSheets() " + e.getMessage());
        }
    }


    private String compareRows(Map<String, List<String>> previousSheetMap,
                               Map<String, List<String>> nextSheetMap,
                               Sheet previousSheet,
                               Sheet nextSheet) throws Exception {

        try {
            // StringBuilder to hold text summary of differences
            StringBuilder differencesToDisplayAsText = new StringBuilder();

            // Set cell color style for the sheet
            CellStyle addedStyle   = createColorCellStyle(nextSheet.getWorkbook(), IndexedColors.GREEN );
            CellStyle removedStyle = createColorCellStyle(nextSheet.getWorkbook(), IndexedColors.RED   );
            CellStyle changedStyle = createColorCellStyle(nextSheet.getWorkbook(), IndexedColors.YELLOW);

            // Merge all keys of each tab in one list to get the whole list of keys
            Set<String> allRowKeys = new LinkedHashSet<>(previousSheetMap.keySet());
            allRowKeys.addAll(nextSheetMap.keySet());

            for (String rowKey : allRowKeys) {
                List<String> rowValuesFromPreviousSheet = previousSheetMap.get(rowKey);
                List<String> rowValuesFromNextSheet     = nextSheetMap.get(rowKey);

                // Check if row is added
                if (rowValuesFromPreviousSheet == null) {
                    this.updateRow(nextSheet, rowKey, addedStyle, "Rangée ajoutée");
                    differencesToDisplayAsText.append("Rangée ajoutée: ")
                                              .append(rowValuesFromNextSheet)
                                              .append("\n");
                }
                // Check if row is deleted
                else if (rowValuesFromNextSheet == null) {
                    this.updateRow(previousSheet, rowKey, removedStyle, "Rangée effacée");
                    differencesToDisplayAsText.append("Rangée effacée: ")
                                              .append(rowValuesFromPreviousSheet)
                                              .append("\n");
                } else {
                    // Check if any value of the row has been modified
                    for (int columnIndex = 0; columnIndex < Math.max(rowValuesFromPreviousSheet.size(), rowValuesFromNextSheet.size()); columnIndex++) {
                        String previousSheetValue = columnIndex < rowValuesFromPreviousSheet.size() ? rowValuesFromPreviousSheet.get(columnIndex) : ""; // Check if the column has been deleted
                        String nextSheetValue = columnIndex < rowValuesFromNextSheet.size() ? rowValuesFromNextSheet.get(columnIndex) : "";         // Check if the column has been deleted
                        // If the cell has been modified
                        if (!previousSheetValue.equals(nextSheetValue)) {
                            updateCell(nextSheet, rowKey, columnIndex + 1, changedStyle); // Mark the cell with color highlight and label
                            differencesToDisplayAsText.append("Colonne ") // Prepare a text to summarize change details
                                                      .append(columnIndex + 1)
                                                      .append(" changée de '")
                                                      .append(previousSheetValue)
                                                      .append("' à '")
                                                      .append(nextSheetValue)
                                                      .append("'.\n");
                        }
                    }
                }
            }
            return differencesToDisplayAsText.toString();
        } catch(Exception e) {
            throw new Exception("Exception from compareRow() " + e.getMessage());
        }
    }



    private List<Map<String, List<String>>> extractSheetsDataFromFile(MultipartFile file) throws Exception {
        List<Map<String, List<String>>> excelSheets = new ArrayList<>();

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {
            for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
                Map<String, List<String>> sheetData = new LinkedHashMap<>();
                Sheet sheet = workbook.getSheetAt(sheetIndex);
                for (Row row : sheet) { // Loop over rows
                    Cell keyCell = row.getCell(0);
                    if (keyCell != null) {
                        String key = convertCellValueAsString(keyCell);
                        List<String> values = new ArrayList<>();
                        for (int cellIndex = 1; cellIndex < row.getLastCellNum(); cellIndex++) { // Loop over cells
                            Cell valueCell = row.getCell(cellIndex);
                            values.add(valueCell != null ? convertCellValueAsString(valueCell) : "");
                        }
                        sheetData.put(key, values);
                    }
                }
                excelSheets.add(sheetData);
            }
            return excelSheets;
        } catch(Exception e) {
            throw new Exception("Exception from extractSheetsDataFromFile() " + e.getMessage());
        }
    }


    private void updateRow(Sheet sheet, String key, CellStyle style, String status) throws Exception {
        try {
            for (Row row : sheet) { // Parse all rows of the sheet
                Cell keyCell = row.getCell(0); // Get key/first cell of the row
                if (keyCell != null && key.equals(convertCellValueAsString(keyCell))) { // Control we're highlighting the right value
                    for (Cell cell : row) {
                        cell.setCellStyle(style);
                    }
                    keyCell.setCellValue(keyCell.getStringCellValue() + " -> " + status);
                    break;
                }
            }
        } catch(Exception e) {
            throw new Exception("Exception from updateRow " + e.getMessage());
        }
    }


    private void updateCell(Sheet sheet, String key, int columnIndex, CellStyle style) throws Exception {
        try {
            for (Row row : sheet) {
                Cell keyCell = row.getCell(0);
                if (keyCell != null && key.equals(convertCellValueAsString(keyCell))) {
                    Cell cell = row.getCell(columnIndex);
                    if (cell == null) {
                        cell = row.createCell(columnIndex);
                    }
                    cell.setCellStyle(style);
                    cell.setCellValue(this.convertCellValueAsString(cell) + " -> Cellule modifiée");
                    break;
                }
            }
        } catch(Exception e) {
            throw new Exception("Exception from updateCell " + e.getMessage());
        }
    }


    private CellStyle createColorCellStyle(Workbook workbook, IndexedColors color) throws Exception {
        try {
            CellStyle style = workbook.createCellStyle();
            style.setFillForegroundColor(color.getIndex());
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            return style;
        } catch(Exception e) {
            throw new Exception("Exception from createColorCellStyle " + e.getMessage());
        }
    }


    private String convertCellValueAsString(Cell cell) throws Exception {
        try {
            switch (cell.getCellType()) {
                case STRING:
                    return cell.getStringCellValue();
                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        return cell.getDateCellValue().toString();
                    } else {
                        double numericValue = cell.getNumericCellValue();
                        if (numericValue == (int) numericValue) {
                            return String.valueOf((int) numericValue);
                        } else {
                            return String.valueOf(numericValue);
                        }
                    }
                case BOOLEAN:
                    return String.valueOf(cell.getBooleanCellValue());
                case FORMULA:
                    return cell.getCellFormula();
                case BLANK:
                    return "";
                default:
                    return cell.toString();
            }
        } catch(Exception e) {
            throw new Exception("Exception from convertCellValueAsString " + e.getMessage());
        }
    }


}
