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

        headerRow.createCell(30).setCellValue("snknode_req_sys_trace");
        headerRow.createCell(31).setCellValue("SYS_TRACE_AUDIT_NO");

        headerRow.createCell(32).setCellValue("snknode_rev_sys_trace");
        headerRow.createCell(33).setCellValue("SYS_TRACE_AUDIT_NO");

        headerRow.createCell(35).setCellValue("TRAN_DESC");
        headerRow.createCell(36).setCellValue("SYS_TRACE_AUDIT_NO");

        headerRow.createCell(37).setCellValue("snknode_date_settle");
        headerRow.createCell(38).setCellValue("DATE_RECON_ISS");

        headerRow.createCell(39).setCellValue("DATE_RECON_NET");

        headerRow.createCell(40).setCellValue("snknode_amount_requested");
        headerRow.createCell(41).setCellValue("AMT_RECON_ISS");
        headerRow.createCell(42).setCellValue("AMT_RECON_NET");
        headerRow.createCell(43).setCellValue("O_AMT_CARD_BILL");

        headerRow.createCell(44).setCellValue("O_AMT_RECON_ISS");
        headerRow.createCell(45).setCellValue("O_AMT_RECON_NET");
        headerRow.createCell(46).setCellValue("O_AMT_RECON_ISS");

        headerRow.createCell(47).setCellValue("snknode_cash_requested");
        headerRow.createCell(48).setCellValue("ADL_RESP_AMT1");

        headerRow.createCell(49).setCellValue("snknode_currency_code");
        headerRow.createCell(50).setCellValue("CUR_RECON_ISS");
        headerRow.createCell(51).setCellValue("CUR_RECON_NET");
        headerRow.createCell(52).setCellValue("CUR_CARD_BILL");

        headerRow.createCell(53).setCellValue("snknode_conversion_rate");
        headerRow.createCell(54).setCellValue("CNV_RCN_ISS_DE_POS");
        headerRow.createCell(55).setCellValue("CNV_RCN_ISS_RATE");

        headerRow.createCell(56).setCellValue("snknode_date_conversion");
        headerRow.createCell(57).setCellValue("DATE_CVV_ISS");

        headerRow.createCell(58).setCellValue("totals_group");
        headerRow.createCell(59).setCellValue("INST_ID_RECON_ISS");
        headerRow.createCell(60).setCellValue("INST_ID_RECN_ISS_B");

        headerRow.createCell(61).setCellValue("tran_type");
        headerRow.createCell(62).setCellValue("TRAN_TYPE_ID");

        headerRow.createCell(63).setCellValue("from_account");
        headerRow.createCell(64).setCellValue("TRAN_TYPE_ID");

        headerRow.createCell(65).setCellValue("to_account");
        headerRow.createCell(66).setCellValue("TRAN_TYPE_ID");

        headerRow.createCell(67).setCellValue("amount_tran_requested");
        headerRow.createCell(68).setCellValue("AMT_TRAN");
        headerRow.createCell(69).setCellValue("O_AMT_TRAN");

        headerRow.createCell(70).setCellValue("amount_cash_requested");
        headerRow.createCell(71).setCellValue("ADL_RQST_AMT0");

        headerRow.createCell(72).setCellValue("gmt_date_time");
        headerRow.createCell(73).setCellValue("GMT_TIME");

        headerRow.createCell(74).setCellValue("time_local");
        headerRow.createCell(75).setCellValue("TSTAMP_LOCAL");

        headerRow.createCell(76).setCellValue("acct_no");
        headerRow.createCell(77).setCellValue("TSTAMP_LOCAL");

        headerRow.createCell(78).setCellValue("expiry_dat");
        headerRow.createCell(79).setCellValue("DATE_EXP");

        headerRow.createCell(80).setCellValue("merchant_type");
        headerRow.createCell(81).setCellValue("MERCH_TYPE");


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
        outputRow.createCell(5).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get("TRAN_DESC")))); // TRAN_DESC from Realtime_Data
        outputRow.createCell(6).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get("DRAFT_CAPTURE_FLG")))); // DRAFT_CAPTURE_FLG from DN_Data
        outputRow.createCell(7).setCellValue(getCellValueAsString(rtDataRow.getCell(rtColumnMap.get("draft_capture")))); // DRAFT_CAPTURE_FLG_RT from Realtime_Data
        outputRow.createCell(8).setCellValue(getCellValueAsString(rtDataRow.getCell(rtColumnMap.get("stan_in")))); // stan_in from Realtime_Data
        outputRow.createCell(9).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get("STANDIN_ACT")))); // STANDIN_ACT from DN_Data
        outputRow.createCell(10).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get("DATE_RECON_ACQ")))); // DATE_RECON_ACQ from DN_Data
        outputRow.createCell(11).setCellValue(getCellValueAsString(rtDataRow.getCell(rtColumnMap.get("srcnode_date_settle")))); // srcnode_date_settle from Realtime_Data
        outputRow.createCell(12).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get("ADL_RQST_AMT1")))); // ADL_RQST_AMT1 from DN_Data
        outputRow.createCell(13).setCellValue(getCellValueAsString(rtDataRow.getCell(rtColumnMap.get("srcnode_cash_requested")))); // srcnode_cash_requested from Realtime_Data
        outputRow.createCell(14).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get("CUR_RECON_ACQ_1")))); // CUR_RECON_ACQ_1 from DN_Data
        outputRow.createCell(15).setCellValue(getCellValueAsString(rtDataRow.getCell(rtColumnMap.get("srcnode_currency_code")))); // srcnode_currency_code from Realtime_Data
        outputRow.createCell(16).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get("CNV_RCN_ACQ_RATE")))); // CNV_RCN_ACQ_RATE from DN_Data
        outputRow.createCell(17).setCellValue(getCellValueAsString(rtDataRow.getCell(rtColumnMap.get("srcnode_conversion_rate")))); // srcnode_conversion_rate from Realtime_Data
        outputRow.createCell(18).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get("CNV_RCN_ACQ_DE_POS")))); // CNV_RCN_ACQ_DE_POS from DN_Data
        outputRow.createCell(19).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get("DATE_CNV_ACQ")))); // DATE_CNV_ACQ from Realtime_Data
        outputRow.createCell(20).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get("ODE_MTI")))); // ODE_MTI from DN_Data
        outputRow.createCell(21).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get("ODE_SYS_TRA_AUD_NO")))); // ODE_SYS_TRA_AUD_NO from DN_Data
        outputRow.createCell(22).setCellValue(getCellValueAsString(rtDataRow.getCell(rtColumnMap.get("srcnode_original_data")))); // srcnode_original_data from Realtime_Data
        outputRow.createCell(23).setCellValue(getCellValueAsString(rtDataRow.getCell(rtColumnMap.get("snknode_req_sys_trace")))); // snknode_req_sys_trace from Realtime_Data
        outputRow.createCell(24).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get("SYS_TRACE_AUDIT_NO")))); // SYS_TRACE_AUDIT_NO from DN_Data
        outputRow.createCell(25).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get("TRAN_DESC")))); // TRAN_DESC from DN_Data
        outputRow.createCell(26).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get("DATE_RECON_ISS")))); // DATE_RECON_ISS from Realtime_Data
        outputRow.createCell(27).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get("DATE_RECON_NET")))); // DATE_RECON_NET from Realtime_Data
        outputRow.createCell(28).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get("AMT_RECON_ISS")))); // AMT_RECON_ISS from DN_Data
        outputRow.createCell(29).setCellValue(getCellValueAsString(rtDataRow.getCell(dnColumnMap.get("AMT_RECON_NET")))); // AMT_RECON_NET from Realtime_Data
        outputRow.createCell(30).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get("O_AMT_CARD_BILL")))); // O_AMT_CARD_BILL from DN_Data
       // outputRow.createCell(31).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get("O_AMT_RECON_ISS")))); // O_AMT_RECON_ISS from DN_Data
       // outputRow.createCell(32).setCellValue(getCellValueAsString(rtDataRow.getCell(rtColumnMap.get("snknode_cash_requested")))); // snknode_cash_requested from Realtime_Data
       // outputRow.createCell(33).setCellValue(getCellValueAsString(rtDataRow.getCell(rtColumnMap.get("ADL_RESP_AMT1")))); // ADL_RESP_AMT1 from Realtime_Data
        outputRow.createCell(34).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get("CUR_RECON_ISS")))); // CUR_RECON_ISS from DN_Data
        outputRow.createCell(35).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get("CUR_RECON_NET")))); // CUR_RECON_NET from Realtime_Data
        outputRow.createCell(36).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get("CUR_CARD_BILL")))); // CUR_CARD_BILL from Realtime_Data
        outputRow.createCell(37).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get("CNV_RCN_ISS_RATE")))); // CNV_RCN_ISS_RATE from Realtime_Data
        outputRow.createCell(38).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get("DATE_CNV_ISS")))); // DATE_CNV_ISS from Realtime_Data
        outputRow.createCell(39).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get("INST_ID_RECON_ISS")))); // INST_ID_RECON_ISS from DN_Data
        outputRow.createCell(40).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get("TRAN_TYPE_ID")))); // TRAN_TYPE_ID from Realtime_Data
        outputRow.createCell(41).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get("AMT_TRAN")))); // AMT_TRAN from Realtime_Data
        outputRow.createCell(42).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get("O_AMT_TRAN")))); // O_AMT_TRAN from Realtime_Data
        outputRow.createCell(43).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get("GMT_TIME")))); // GMT_TIME from Realtime_Data
        outputRow.createCell(44).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get("TSTAMP_LOCAL")))); // TSTAMP_LOCAL from Realtime_Data
        outputRow.createCell(45).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get("DATE_EXP")))); // DATE_EXP from Realtime_Data
        outputRow.createCell(46).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get("MERCH_TYPE")))); // MERCH_TYPE from Realtime_Data
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
