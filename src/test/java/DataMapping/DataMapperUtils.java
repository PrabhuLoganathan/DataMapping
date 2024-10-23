package DataMapping;


import org.apache.poi.ss.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class DataMapperUtils {

    // Creates a map of transaction numbers to rows from the DN_Data sheet
    public static Map<String, Row> createDataMap(Sheet sheet, Map<String, Integer> columnMap) {
        Map<String, Row> dataMap = new HashMap<>();
        int tranNumberColumn = columnMap.get("TRAN_NUMBER");

        for (Row row : sheet) {
            String tranNumber = getCellValueAsString(row.getCell(tranNumberColumn));
            if (tranNumber != null && !tranNumber.isBlank()) {
                dataMap.put(tranNumber, row);
            }
        }
        return dataMap;
    }

    // Creates the header row for the output sheet
    public static void createHeaderRow(Sheet outputSheet) {
        Row headerRow = outputSheet.createRow(0);
        headerRow.createCell(0).setCellValue("msg_type");
        headerRow.createCell(1).setCellValue("MTI");
        headerRow.createCell(2).setCellValue("FUNC_CODE");
        headerRow.createCell(3).setCellValue("TRAN_NUMBER_DN_Data");
        headerRow.createCell(4).setCellValue("TRAN_NUMBER_Realtime_Data");
        headerRow.createCell(5).setCellValue("SUBSTR(R.TRAN_DESC,13)");
        headerRow.createCell(6).setCellValue("DRAFT_CAPTURE_FLG_DN");
        headerRow.createCell(7).setCellValue("DRAFT_CAPTURE_FLG_RT");
        headerRow.createCell(8).setCellValue("stan_in");
        headerRow.createCell(9).setCellValue("STANDIN_ACT");
        headerRow.createCell(10).setCellValue("DATE_RECON_ACQ");
        headerRow.createCell(11).setCellValue("srcnode_date_settle");
        headerRow.createCell(12).setCellValue("ADL_RQST_AMT1");
        headerRow.createCell(13).setCellValue("srcnode_cash_requested");
        headerRow.createCell(14).setCellValue("CUR_RECON_ACQ_1");
        headerRow.createCell(15).setCellValue("srcnode_currency_code");

        headerRow.createCell(16).setCellValue("CNV_RCN_ACQ_RATE");
        headerRow.createCell(17).setCellValue("srcnode_conversion_rate");

        headerRow.createCell(18).setCellValue("CNV_RCN_ACQ_DE_POS");
        headerRow.createCell(19).setCellValue("srcnode_conversion_rate");

        headerRow.createCell(20).setCellValue("DATE_CNV_ACQ");
        headerRow.createCell(21).setCellValue("srcnode_date_conversion");

        headerRow.createCell(22).setCellValue("ODE_MTI");
        headerRow.createCell(23).setCellValue("srcnode_original_data");

        headerRow.createCell(24).setCellValue("ODE_SYS_TRA_AUD_NO");
        headerRow.createCell(25).setCellValue("srcnode_original_data");

        headerRow.createCell(26).setCellValue("ODE_TSTAMP_LOCL_TR");
        headerRow.createCell(27).setCellValue("srcnode_original_data");

        headerRow.createCell(26).setCellValue("ODE_INST_ID_ACQ");
        headerRow.createCell(27).setCellValue("srcnode_original_data");


    }

    // Processes the Realtime_Data sheet, matching transactions with DN_Data and writing results to the output sheet
    public static void processRealtimeSheet(Sheet rtSheet, Map<String, Row> dnDataMap, Sheet outputMappingSheet, Map<String, Integer> rtColumnMap, Map<String, Integer> dnColumnMap) {
        int outputRowNum = 1;
        for (Row row : rtSheet) {
            String tranNumberRealtime = getCellValueAsString(row.getCell(rtColumnMap.get("tran_nr")));
            if (dnDataMap.containsKey(tranNumberRealtime)) {
                Row outputRow = outputMappingSheet.createRow(outputRowNum++);
                Row dnDataRow = dnDataMap.get(tranNumberRealtime);

                // Populate the output row with relevant data from both sheets
                mapRowData(outputRow, dnDataRow, row, rtColumnMap, dnColumnMap);
            }
        }
    }

    // Maps the data from DN and Realtime rows into the output row using column names
    private static void mapRowData(Row outputRow, Row dnDataRow, Row rtDataRow, Map<String, Integer> rtColumnMap, Map<String, Integer> dnColumnMap) {
        outputRow.createCell(0).setCellValue(getCellValueAsString(rtDataRow.getCell(rtColumnMap.get("msg_type")))); // msg_type from Realtime_Data
        outputRow.createCell(1).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get("MTI")))); // MTI from DN_Data
        outputRow.createCell(2).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get("FUNC_CODE")))); // FUNC_CODE from DN_Data
        outputRow.createCell(3).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get("TRAN_NUMBER")))); // TRAN_NUMBER from DN_Data
        outputRow.createCell(4).setCellValue(getCellValueAsString(rtDataRow.getCell(rtColumnMap.get("tran_nr")))); // tran_nr from Realtime_Data
        outputRow.createCell(5).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get("SUBSTR(R.TRAN_DESC,13)")))); // TRAN_DESC_SUBSTR
        outputRow.createCell(6).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get("DRAFT_CAPTURE_FLG")))); // Draft Capture from DN_Data
        outputRow.createCell(7).setCellValue(getCellValueAsString(rtDataRow.getCell(rtColumnMap.get("draft_capture")))); // Draft Capture from Realtime_Data
        outputRow.createCell(8).setCellValue(getCellValueAsString(rtDataRow.getCell(rtColumnMap.get("stan_in")))); // stan_in from Realtime_Data
        outputRow.createCell(9).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get("STANDIN_ACT")))); // STANDIN_ACT from DN_Data
        outputRow.createCell(10).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get("DATE_RECON_ACQ")))); // DATE_RECON_ACQ from DN_Data
        outputRow.createCell(11).setCellValue(getCellValueAsString(rtDataRow.getCell(rtColumnMap.get("srcnode_date_settle")))); // srcnode_date_settle from Realtime_Data
        outputRow.createCell(12).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get("ADL_RQST_AMT1")))); // ADL_RQST_AMT1 from DN_Data
        outputRow.createCell(13).setCellValue(getCellValueAsString(rtDataRow.getCell(rtColumnMap.get("srcnode_cash_requested")))); // srcnode_cash_requested from Realtime_Data
        outputRow.createCell(14).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get("CUR_RECON_ACQ_1")))); // CUR_RECON_ACQ_1 from DN_Data
        outputRow.createCell(15).setCellValue(getCellValueAsString(rtDataRow.getCell(rtColumnMap.get("srcnode_currency_code")))); // srcnode_currency_code from Realtime_Data
        outputRow.createCell(16).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get("CNV_RCN_ACQ_RATE")))); // CNV_RCN_ACQ_RATE" from DN_Data
        outputRow.createCell(17).setCellValue(getCellValueAsString(rtDataRow.getCell(rtColumnMap.get("srcnode_conversion_rate")))); // srcnode_conversion_rate from Realtime_Data
        outputRow.createCell(18).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get("CNV_RCN_ACQ_DE_POS")))); // CNV_RCN_ACQ_RATE" from DN_Data
        outputRow.createCell(19).setCellValue(getCellValueAsString(rtDataRow.getCell(rtColumnMap.get("srcnode_conversion_rate")))); // srcnode_conversion_rate from Realtime_Data

    }

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

    // Method to create a map of column names to their respective column indexes
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
}
