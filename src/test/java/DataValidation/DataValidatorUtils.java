package DataValidation;


import org.apache.poi.ss.usermodel.*;

import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Arrays;

public class DataValidatorUtils {

    // Map to support multiple msgType values per MTI
    private static final Map<String, List<Integer>> validMtiMsgTypeMap = new HashMap<>();
    private static final Map<String, Integer> funcCodeMap = new HashMap<>();

    // Initialize valid MTI and message type mappings
    public static void initializeValidMtiMsgTypeMap() {
        validMtiMsgTypeMap.put("1430", Arrays.asList(1057, 1056));
        validMtiMsgTypeMap.put("1110", Arrays.asList(256));
        validMtiMsgTypeMap.put("1210", Arrays.asList(512));
        validMtiMsgTypeMap.put("1230", Arrays.asList(544, 545));

        funcCodeMap.put("1210", 200);
        funcCodeMap.put("1230", 200);
        funcCodeMap.put("1110", 100);
        funcCodeMap.put("1430", 400);
        funcCodeMap.put("1410", 400);
    }

    // Get the map of column names to their indexes
    public static Map<String, Integer> getColumnMap(Row headerRow) {
        Map<String, Integer> columnMap = new HashMap<>();
        for (Cell cell : headerRow) {
            columnMap.put(cell.getStringCellValue(), cell.getColumnIndex());
        }
        return columnMap;
    }

    // Copy row values to output row
    public static void copyRowValues(Row outputRow, Row inputRow) {
        for (int i = 0; i < inputRow.getPhysicalNumberOfCells(); i++) {
            Cell inputCell = inputRow.getCell(i);
            Cell outputCell = outputRow.createCell(i);

            if (inputCell != null) {
                switch (inputCell.getCellType()) {
                    case STRING:
                        outputCell.setCellValue(inputCell.getStringCellValue());
                        break;
                    case NUMERIC:
                        outputCell.setCellValue(inputCell.getNumericCellValue());
                        break;
                    case BOOLEAN:
                        outputCell.setCellValue(inputCell.getBooleanCellValue());
                        break;
                    case FORMULA:
                        outputCell.setCellFormula(inputCell.getCellFormula());
                        break;
                    default:
                        outputCell.setCellValue("");
                        break;
                }
            }
        }
    }

    // Copy header row and add validation columns
    public static void copyHeaderRow(Sheet outputSheet, Row headerRow) {
        Row outputHeaderRow = outputSheet.createRow(0);
        for (int i = 0; i < headerRow.getPhysicalNumberOfCells(); i++) {
            Cell inputCell = headerRow.getCell(i);
            Cell outputCell = outputHeaderRow.createCell(i);

            if (inputCell != null) {
                //outputCell.setCellValue(inputCell.getStringCellValue());
            }
        }

        // Adding new header columns for validation results
        //int colIndex = headerRow.getPhysicalNumberOfCells();
        int colIndex = 0;

        outputHeaderRow.createCell(colIndex++).setCellValue("msg_type and MTI validation");
        outputHeaderRow.createCell(colIndex++).setCellValue("Function Code");
        outputHeaderRow.createCell(colIndex++).setCellValue("Draft Capture Validation");
        outputHeaderRow.createCell(colIndex++).setCellValue("Validate_StanIn_StandinAct");
        outputHeaderRow.createCell(colIndex++).setCellValue("Validate_DATE_RECON_ACQ_srcnode_date_settle");
        outputHeaderRow.createCell(colIndex++).setCellValue("Validate_ADL_RQST_AMT1_srcnode_cash_requested");
        outputHeaderRow.createCell(colIndex++).setCellValue("Validate_CUR_RECON_ACQ_1_srcnode_currency_code");
    }

    // Validation methods
    public static boolean isValidMtiMsgType(String msgType, String mti) {
        if (msgType == null || msgType.isEmpty() || mti == null || mti.isEmpty()) {
            return false;
        }

        try {
            int msgTypeValue = Integer.parseInt(msgType);
            if (validMtiMsgTypeMap.containsKey(mti)) {
                List<Integer> expectedMsgTypes = validMtiMsgTypeMap.get(mti);
                return expectedMsgTypes.contains(msgTypeValue);
            }
        } catch (NumberFormatException e) {
            System.out.println("Error parsing msgType or MTI: " + msgType + ", " + mti);
        }

        return false;
    }

    // Validate draft capture flags for both DN and RT
    public static boolean isValidDraftCaptureFlags(String draftFlgDn, String draftFlgRt) {
        // Assuming valid values are 0 or 1
        return (draftFlgDn.equals("0") || draftFlgDn.equals("1")) && (draftFlgRt.equals("0") || draftFlgRt.equals("1"));
    }

    // Method to fetch function code based on MTI
    public static String getFunctionCode(String mti) {
        return funcCodeMap.containsKey(mti) ? String.valueOf(funcCodeMap.get(mti)) : "Not Found";
    }

    // Method to compare stan_in and STANDIN_ACT
    public static boolean compareStanInAndStandinAct(String stanIn, String standinAct) {
        return stanIn.equals(standinAct);
    }


    // Method to compare sourceNodeCashRequested and additionalRequestAmount
    public static boolean compareSourceNodeCashRequestedAndAdditionalRequestAmount(String additionalRequestAmount, String sourceNodeCashRequested) {
        return additionalRequestAmount.equals(sourceNodeCashRequested);
    }

    // Method to compare CUR_RECON_ACQ_1 and srcnode_currency_code
    public static boolean compareCurReconAcqAndSrcnodeCurrencyCode(String curReconAcq1, String srcnodeCurrencyCode) {
        return curReconAcq1.equals(srcnodeCurrencyCode);
    }


    // Method to compare DATE_RECON_ACQ (YYYYMMDD) and srcnode_date_settle (MMDD)
    public static boolean compareDatesWithYearIgnored(String dateReconAcq, String srcnodeDateSettle) {
        if (dateReconAcq.length() >= 8 && srcnodeDateSettle.length() == 3) {
            String dateReconAcqMMDD = dateReconAcq.substring(5); // Get MMDD from YYYYMMDD
            System.out.println(dateReconAcqMMDD);
            System.out.println(srcnodeDateSettle);
            return dateReconAcqMMDD.equals(srcnodeDateSettle); // Compare MMDD
        }
        return false;
    }

    // Helper method to get cell value as String
    public static String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf((int) cell.getNumericCellValue());
            default:
                return "";
        }
    }
}
