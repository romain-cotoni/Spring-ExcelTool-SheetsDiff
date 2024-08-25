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
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayOutputStream;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

@Service
public class ExcelToolsService {

    public ByteArrayOutputStream findDifferencesBetweenSheets(MultipartFile file) throws Exception {
        try {
            Workbook workbook = new XSSFWorkbook(file.getInputStream());

            ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
            if(workbook.getNumberOfSheets() > 1) {

                // Extract data from file before updating
                List<Map<String, List<String>>> excelSheetsMap = this.mapRowKeyValuesFromSheets(file);

                // Parse all sheets
                for (int sheetIndex = 1; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
                    compareSheets(workbook.getSheetAt(sheetIndex - 1),
                                  workbook.getSheetAt(sheetIndex),
                                  excelSheetsMap.get(sheetIndex - 1),
                                  excelSheetsMap.get(sheetIndex));
                }

                // Save into byteArrayOutputStream an updated Excel file with highlighted and labeled differences
                workbook.write(byteArrayOutputStream);
                workbook.close();
            }
            return byteArrayOutputStream;
        } catch (Exception e) {
            throw new Exception("Exception from findDifferencesBetweenSheets() " + e.getMessage());
        }
    }


    private void compareSheets(Sheet previousSheet,
                               Sheet nextSheet,
                               Map<String, List<String>> previousSheetMap,
                               Map<String, List<String>> nextSheetMap) throws Exception {
        try {
            // Set cell color style for the sheet
            CellStyle addedStyle   = createColorCellStyle(nextSheet.getWorkbook(), IndexedColors.LIGHT_GREEN );
            CellStyle removedStyle = createColorCellStyle(nextSheet.getWorkbook(), IndexedColors.LIGHT_TURQUOISE);
            CellStyle changedStyle = createColorCellStyle(nextSheet.getWorkbook(), IndexedColors.LIGHT_YELLOW);

            // Since rows can be deleted or added, we merge all keys of each sheets in one list in order to get the whole list of distinct keys
            Set<String> allKeys = new LinkedHashSet<>(previousSheetMap.keySet());
            allKeys.addAll(nextSheetMap.keySet());

            // Parse all rows
            for (String key : allKeys) {
                List<String> previousSheetRowValues = previousSheetMap.get(key);
                List<String> nextSheetRowValues     = nextSheetMap.get(key);

                // Check if row is added
                if (previousSheetRowValues == null) {
                    this.updateRow(nextSheet, key, addedStyle, "Rangée ajoutée");
                }
                // Check if row is deleted
                else if (nextSheetRowValues == null) {
                    this.updateRow(previousSheet, key, removedStyle, "Rangée effacée");
                } else {
                    // Check if any value of the row has been modified
                    this.compareRows(previousSheetRowValues, nextSheetRowValues, nextSheet, key, changedStyle);
                }
            }
        } catch(Exception e) {
            throw new Exception("Exception from compareSheets() " + e.getMessage());
        }
    }


    private void compareRows(List<String> previousSheetRowValues,
                             List<String> nextSheetRowValues,
                             Sheet nextSheet,
                             String key,
                             CellStyle changedStyle ) throws Exception {
        try {
            for (int columnIndex = 0; columnIndex < Math.max(previousSheetRowValues.size(),nextSheetRowValues.size()); columnIndex++) {

                // Check if the column has been deleted
                String previousSheetValue = columnIndex < previousSheetRowValues.size() ? previousSheetRowValues.get(columnIndex) : "";
                String nextSheetValue     = columnIndex < nextSheetRowValues.size()     ? nextSheetRowValues.get(columnIndex)     : "";

                // If the cell has been modified
                if (!previousSheetValue.equals(nextSheetValue)) {
                    // Mark the cell with color highlight and label
                    updateCell(nextSheet, key, columnIndex + 1, changedStyle);
                }
            }
        } catch(Exception e) {
            throw new Exception("Exception from compareRows() " + e.getMessage());
        }
    }


    private List<Map<String, List<String>>> mapRowKeyValuesFromSheets(MultipartFile file) throws Exception {
        List<Map<String, List<String>>> sheetsList = new ArrayList<>();

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {
            for(int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
                Map<String, List<String>> sheetMap = new LinkedHashMap<>();
                for (Row row : workbook.getSheetAt(sheetIndex)) { // Loop over rows
                    Cell keyCell = row.getCell(0);
                    if (keyCell != null) {
                        String key = convertCellValueAsString(keyCell);
                        List<String> values = new ArrayList<>();
                        for (int cellIndex = 1; cellIndex < row.getLastCellNum(); cellIndex++) { // Loop over cells
                            Cell valueCell = row.getCell(cellIndex);
                            values.add(valueCell != null ? convertCellValueAsString(valueCell) : "");
                        }
                        sheetMap.put(key, values);
                    }
                }
                sheetsList.add(sheetMap);
            }
            return sheetsList;
        } catch(Exception e) {
            throw new Exception("Exception from mapRowKeyValuesFromSheets() " + e.getMessage());
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
