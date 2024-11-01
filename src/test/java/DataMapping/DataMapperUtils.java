package DataMapping;

import org.apache.poi.ss.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class DataMapperUtils {

    // Helper method to get the value of a cell as a String
    public static String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    return String.valueOf((long) cell.getNumericCellValue());
                }
            case BOOLEAN:
                return Boolean.toString(cell.getBooleanCellValue());
            case FORMULA:
                FormulaEvaluator evaluator = cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
                CellValue cellValue = evaluator.evaluate(cell);
                switch (cellValue.getCellType()) {
                    case STRING:
                        return cellValue.getStringValue();
                    case NUMERIC:
                        return String.valueOf((long) cellValue.getNumberValue());
                    case BOOLEAN:
                        return Boolean.toString(cellValue.getBooleanValue());
                    default:
                        return "";
                }
            default:
                return "";
        }
    }



    // Copy data from one sheet to another
    public static void copySheetToOutput(Workbook outputWorkbook, Sheet sourceSheet, String sheetName) {
        Sheet outputSheet = outputWorkbook.createSheet(sheetName);
        for (int i = 0; i < sourceSheet.getPhysicalNumberOfRows(); i++) {
            Row sourceRow = sourceSheet.getRow(i);
            Row destinationRow = outputSheet.createRow(i);
            if (sourceRow != null) {
                copyRowData(sourceRow, destinationRow);
            }
        }
    }

    // Copies data from one row to another
    private static void copyRowData(Row sourceRow, Row destinationRow) {
        for (int j = 0; j < sourceRow.getPhysicalNumberOfCells(); j++) {
            Cell sourceCell = sourceRow.getCell(j);
            Cell destinationCell = destinationRow.createCell(j);
            if (sourceCell != null) {
                switch (sourceCell.getCellType()) {
                    case STRING:
                        destinationCell.setCellValue(sourceCell.getStringCellValue());
                        break;
                    case NUMERIC:
                        destinationCell.setCellValue(sourceCell.getNumericCellValue());
                        break;
                    case BOOLEAN:
                        destinationCell.setCellValue(sourceCell.getBooleanCellValue());
                        break;
                    case FORMULA:
                        destinationCell.setCellFormula(sourceCell.getCellFormula());
                        break;
                    default:
                        destinationCell.setCellValue("");

                }
            }
        }
    }

    // Writes the output workbook to the specified file path
    public static void writeOutputFile(Workbook workbook, String outputPath) {
        try (FileOutputStream outputStream = new FileOutputStream(outputPath)) {
            workbook.write(outputStream);
            System.out.println("Output saved to " + outputPath);
        } catch (IOException e) {
            System.out.println("Error while saving the output file.");
            e.printStackTrace();
        }
    }



    public static Map<String, Integer> getColumnMap(Row headerRow) {
        Map<String, Integer> columnMap = new HashMap<>();
        if (headerRow == null) {
            throw new IllegalArgumentException("Header row cannot be null");
        }

        // Iterate through the cells in the header row to build the map
        for (Cell cell : headerRow) {
            String columnName = getCellValueAsString(cell);
            if (columnName != null && !columnName.isEmpty()) {
                columnMap.put(columnName, cell.getColumnIndex());
            }
        }
        return columnMap;
    }
    // Create header row in the output sheet by combining headers from RT and DN sheets
    public static void createCombinedHeaderRow(Sheet outputSheet, Map<String, Integer> rtColumnMap, Map<String, Integer> dnColumnMap) {
        Row headerRow = outputSheet.createRow(0);

        int cellIndex = 0;
        // Add headers from RT sheet
        for (String rtHeader : rtColumnMap.keySet()) {
            headerRow.createCell(cellIndex++).setCellValue(rtHeader);
        }
        // Add headers from DN sheet, skipping any duplicate column names
        for (String dnHeader : dnColumnMap.keySet()) {
            if (!rtColumnMap.containsKey(dnHeader)) {  // Avoid duplicate headers
                headerRow.createCell(cellIndex++).setCellValue(dnHeader);
            }
        }
    }

    // Populate the output sheet with combined data from RT and DN sheets
    public static void populateCombinedOutputSheet(Sheet outputSheet, Map<String, TransactionData> transactionDataMap, Map<String, Integer> rtColumnMap, Map<String, Integer> dnColumnMap) {
        int outputRowNum = 1;

        for (Map.Entry<String, TransactionData> entry : transactionDataMap.entrySet()) {
            String tranNumber = entry.getKey();
            TransactionData transactionData = entry.getValue();
            Row outputRow = outputSheet.createRow(outputRowNum++);

            int cellIndex = 0;

            // Fill RT data columns
            for (String rtHeader : rtColumnMap.keySet()) {
                outputRow.createCell(cellIndex++).setCellValue(transactionData.getRtData(rtHeader));
            }

            // Fill DN data columns, skipping any duplicate headers
            for (String dnHeader : dnColumnMap.keySet()) {
                if (!rtColumnMap.containsKey(dnHeader)) {  // Avoid duplicate headers
                    outputRow.createCell(cellIndex++).setCellValue(transactionData.getDnData(dnHeader));
                }
            }
        }
    }

    // DataMapperUtils.java
    public static void populateTransactionDataMap(Sheet dnSheet, Sheet rtSheet, Map<String, Integer> dnColumnMap, Map<String, Integer> rtColumnMap, Map<String, TransactionData> transactionDataMap) {
        // Populate DN data
        for (int i = 1; i < dnSheet.getPhysicalNumberOfRows(); i++) {  // Start from 1 to skip header row
            Row dnRow = dnSheet.getRow(i);
            String tranNumber = getCellValueAsString(dnRow.getCell(dnColumnMap.get("TRAN_NUMBER")));

            if (tranNumber != null && !tranNumber.isBlank()) {
                TransactionData transactionData = transactionDataMap.getOrDefault(tranNumber, new TransactionData());

                // Loop through DN columns and fill values
                dnColumnMap.forEach((columnName, columnIndex) -> {
                    String cellValue = getCellValueAsString(dnRow.getCell(columnIndex));
                    transactionData.addDnData(columnName, cellValue);
                });

                transactionDataMap.put(tranNumber, transactionData);
            }
        }

        // Populate RT data
        for (int i = 1; i < rtSheet.getPhysicalNumberOfRows(); i++) {  // Start from 1 to skip header row
            Row rtRow = rtSheet.getRow(i);
            String tranNumber = getCellValueAsString(rtRow.getCell(rtColumnMap.get("tran_nr")));

            if (tranNumber != null && !tranNumber.isBlank()) {
                TransactionData transactionData = transactionDataMap.getOrDefault(tranNumber, new TransactionData());

                // Loop through RT columns and fill values
                rtColumnMap.forEach((columnName, columnIndex) -> {
                    String cellValue = getCellValueAsString(rtRow.getCell(columnIndex));
                    transactionData.addRtData(columnName, cellValue);
                });

                transactionDataMap.put(tranNumber, transactionData);
            }
        }
    }


}



