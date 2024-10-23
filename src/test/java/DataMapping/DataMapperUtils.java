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

        // Using RTDataColumns enum for Realtime_Data fields
        headerRow.createCell(0).setCellValue(RTDataColumns.MSG_TYPE.getColumnName());
        headerRow.createCell(1).setCellValue(DNColumns.MTI.getColumnName()); // MTI from DNColumns
        headerRow.createCell(2).setCellValue(DNColumns.FUNC_CODE.getColumnName()); // FUNC_CODE from DNColumns
        headerRow.createCell(3).setCellValue(DNColumns.TRAN_NUMBER.getColumnName() + "_DN_Data"); // TRAN_NUMBER_DN_Data
        headerRow.createCell(4).setCellValue(RTDataColumns.TRAN_NR.getColumnName()); // TRAN_NUMBER from RTDataColumns
        headerRow.createCell(5).setCellValue(DNColumns.SUBSTR_TRAN_DESC.getColumnName()); // SUBSTR(R.TRAN_DESC,13) from DNColumns
        headerRow.createCell(6).setCellValue(DNColumns.DRAFT_CAPTURE_FLG.getColumnName() + "_DN"); // DRAFT_CAPTURE_FLG_DN from DNColumns
        headerRow.createCell(7).setCellValue(RTDataColumns.DRAFT_CAPTURE.getColumnName()); // DRAFT_CAPTURE_FLG_RT from RTDataColumns
        headerRow.createCell(8).setCellValue(RTDataColumns.STAN_IN.getColumnName()); // STAN_IN from RTDataColumns
        headerRow.createCell(9).setCellValue(DNColumns.STANDIN_ACT.getColumnName()); // STANDIN_ACT from DNColumns
        headerRow.createCell(10).setCellValue(DNColumns.DATE_RECON_ACQ.getColumnName()); // DATE_RECON_ACQ from DNColumns
        headerRow.createCell(11).setCellValue(RTDataColumns.SRCNODE_DATE_SETTLE.getColumnName()); // srcnode_date_settle from RTDataColumns
        headerRow.createCell(12).setCellValue(DNColumns.ADL_RQST_AMT1.getColumnName()); // ADL_RQST_AMT1 from DNColumns
        headerRow.createCell(13).setCellValue(RTDataColumns.SRCNODE_CASH_REQUESTED.getColumnName()); // srcnode_cash_requested from RTDataColumns
        headerRow.createCell(14).setCellValue(DNColumns.CUR_RECON_ACQ_1.getColumnName()); // CUR_RECON_ACQ_1 from DNColumns
        headerRow.createCell(15).setCellValue(RTDataColumns.SRCNODE_CURRENCY_CODE.getColumnName()); // srcnode_currency_code from RTDataColumns

        headerRow.createCell(16).setCellValue(DNColumns.CNV_RCN_ACQ_RATE.getColumnName()); // CNV_RCN_ACQ_RATE from DNColumns
        headerRow.createCell(17).setCellValue(RTDataColumns.SRCNODE_CONVERSION_RATE.getColumnName()); // srcnode_conversion_rate from RTDataColumns
        headerRow.createCell(18).setCellValue(DNColumns.CNV_RCN_ACQ_DE_POS.getColumnName()); // CNV_RCN_ACQ_DE_POS from DNColumns
        headerRow.createCell(19).setCellValue(RTDataColumns.SRCNODE_CONVERSION_RATE.getColumnName()); // srcnode_conversion_rate from RTDataColumns

        headerRow.createCell(20).setCellValue(DNColumns.DATE_CNV_ACQ.getColumnName()); // DATE_CNV_ACQ from DNColumns
        headerRow.createCell(21).setCellValue(RTDataColumns.SRCNODE_DATE_CONVERSION.getColumnName()); // srcnode_date_conversion from RTDataColumns

        headerRow.createCell(22).setCellValue(DNColumns.ODE_MTI.getColumnName()); // ODE_MTI from DNColumns
        headerRow.createCell(23).setCellValue(RTDataColumns.SRCNODE_ORIGINAL_DATA.getColumnName()); // srcnode_original_data from RTDataColumns

        headerRow.createCell(24).setCellValue(DNColumns.ODE_SYS_TRA_AUD_NO.getColumnName()); // ODE_SYS_TRA_AUD_NO from DNColumns
        headerRow.createCell(25).setCellValue(RTDataColumns.SRCNODE_ORIGINAL_DATA.getColumnName()); // srcnode_original_data from RTDataColumns

        headerRow.createCell(26).setCellValue(DNColumns.ODE_TSTAMP_LOCAL_TR.getColumnName()); // ODE_TSTAMP_LOCAL_TR from DNColumns
        headerRow.createCell(27).setCellValue(RTDataColumns.SRCNODE_ORIGINAL_DATA.getColumnName()); // srcnode_original_data from RTDataColumns

        headerRow.createCell(30).setCellValue(RTDataColumns.SNKNODE_REQ_SYS_TRACE.getColumnName()); // snknode_req_sys_trace from RTDataColumns
        headerRow.createCell(31).setCellValue(DNColumns.SYS_TRACE_AUDIT_NO.getColumnName()); // SYS_TRACE_AUDIT_NO from DNColumns

        headerRow.createCell(32).setCellValue(RTDataColumns.SNKNODE_REV_SYS_TRACE.getColumnName()); // snknode_rev_sys_trace from RTDataColumns
        headerRow.createCell(33).setCellValue(DNColumns.SYS_TRACE_AUDIT_NO.getColumnName()); // SYS_TRACE_AUDIT_NO from DNColumns

        headerRow.createCell(35).setCellValue(RTDataColumns.SRCNODE_ADDITIONAL_DATA.getColumnName()); // srcnode_additional_data from RTDataColumns
        headerRow.createCell(36).setCellValue(DNColumns.SYS_TRACE_AUDIT_NO.getColumnName()); // SYS_TRACE_AUDIT_NO from DNColumns

        headerRow.createCell(37).setCellValue(RTDataColumns.SNKNODE_DATE_SETTLE.getColumnName()); // snknode_date_settle from RTDataColumns
        headerRow.createCell(38).setCellValue(DNColumns.DATE_RECON_ISS.getColumnName()); // DATE_RECON_ISS from DNColumns

        headerRow.createCell(39).setCellValue(DNColumns.DATE_RECON_NET.getColumnName()); // DATE_RECON_NET from DNColumns

        headerRow.createCell(40).setCellValue(RTDataColumns.SNKNODE_AMOUNT_REQUESTED.getColumnName()); // snknode_amount_requested from RTDataColumns
        headerRow.createCell(41).setCellValue(DNColumns.AMT_RECON_ISS.getColumnName()); // AMT_RECON_ISS from DNColumns
        headerRow.createCell(42).setCellValue(DNColumns.AMT_RECON_NET.getColumnName()); // AMT_RECON_NET from DNColumns
        headerRow.createCell(43).setCellValue(DNColumns.O_AMT_CARD_BILL.getColumnName()); // O_AMT_CARD_BILL from DNColumns

        //headerRow.createCell(44).setCellValue(DNColumns.O_AMT_RECON_ISS.getColumnName()); // O_AMT_RECON_ISS from DNColumns
        headerRow.createCell(45).setCellValue(DNColumns.O_AMT_RECON_NET.getColumnName()); // O_AMT_RECON_NET from DNColumns
       // headerRow.createCell(46).setCellValue(DNColumns.O_AMT_RECON_ISS.getColumnName()); // O_AMT_RECON_ISS from DNColumns

        headerRow.createCell(48).setCellValue(DNColumns.ADL_RESP_AMTO.getColumnName()); // ADL_RESP_AMTO from DNColumns

        headerRow.createCell(49).setCellValue(RTDataColumns.SNKNODE_CURRENCY_CODE.getColumnName()); // snknode_currency_code from RTDataColumns
        headerRow.createCell(50).setCellValue(DNColumns.CUR_RECON_ISS.getColumnName()); // CUR_RECON_ISS from DNColumns
        headerRow.createCell(51).setCellValue(DNColumns.CUR_RECON_NET.getColumnName()); // CUR_RECON_NET from DNColumns
        headerRow.createCell(52).setCellValue(DNColumns.CUR_CARD_BILL.getColumnName()); // CUR_CARD_BILL from DNColumns

        headerRow.createCell(53).setCellValue(RTDataColumns.SNKNODE_CONVERSION_RATE.getColumnName()); // snknode_conversion_rate from RTDataColumns
        headerRow.createCell(54).setCellValue(DNColumns.CNV_RCN_ISS_DE_POS.getColumnName()); // CNV_RCN_ISS_DE_POS from DNColumns
        headerRow.createCell(55).setCellValue(DNColumns.CNV_RCN_ISS_RATE.getColumnName()); // CNV_RCN_ISS_RATE from DNColumns

        headerRow.createCell(56).setCellValue(RTDataColumns.SNKNODE_DATE_CONVERSION.getColumnName()); // snknode_date_conversion from RTDataColumns
        headerRow.createCell(57).setCellValue(DNColumns.DATE_CNV_ISS.getColumnName()); // DATE_CVV_ISS from DNColumns

        headerRow.createCell(58).setCellValue("totals_group");
        headerRow.createCell(59).setCellValue(DNColumns.INST_ID_RECON_ISS.getColumnName()); // INST_ID_RECON_ISS from DNColumns
        headerRow.createCell(60).setCellValue(DNColumns.INST_ID_RECN_ISS_B.getColumnName()); // INST_ID_RECN_ISS_B from DNColumns

        headerRow.createCell(61).setCellValue(RTDataColumns.TRAN_TYPE.getColumnName()); // TRAN_TYPE from RTDataColumns
        headerRow.createCell(62).setCellValue(DNColumns.TRAN_TYPE_ID.getColumnName()); // TRAN_TYPE_ID from DNColumns

        headerRow.createCell(63).setCellValue(RTDataColumns.FROM_ACCOUNT.getColumnName()); // from_account from RTDataColumns
        headerRow.createCell(64).setCellValue(DNColumns.TRAN_TYPE_ID.getColumnName()); // TRAN_TYPE_ID from DNColumns

        headerRow.createCell(65).setCellValue(RTDataColumns.TO_ACCOUNT.getColumnName()); // to_account from RTDataColumns
        headerRow.createCell(66).setCellValue(DNColumns.TRAN_TYPE_ID.getColumnName()); // TRAN_TYPE_ID from DNColumns

        headerRow.createCell(67).setCellValue(RTDataColumns.AMOUNT_TRAN_REQUESTED.getColumnName()); // amount_tran_requested from RTDataColumns
        headerRow.createCell(68).setCellValue(DNColumns.AMT_TRAN.getColumnName()); // AMT_TRAN from DNColumns
        headerRow.createCell(69).setCellValue(DNColumns.O_AMT_TRAN.getColumnName()); // O_AMT_TRAN from DNColumns

        headerRow.createCell(70).setCellValue(RTDataColumns.AMOUNT_CASH_REQUESTED.getColumnName()); // amount_cash_requested from RTDataColumns
        headerRow.createCell(71).setCellValue(DNColumns.ADL_RQST_AMT1.getColumnName()); // ADL_RQST_AMT0 from DNColumns

        headerRow.createCell(72).setCellValue(RTDataColumns.GMT_DATE_TIME.getColumnName()); // gmt_date_time from RTDataColumns
        headerRow.createCell(73).setCellValue(DNColumns.GMT_TIME.getColumnName()); // GMT_TIME from DNColumns

        headerRow.createCell(74).setCellValue(RTDataColumns.TIME_LOCAL.getColumnName()); // time_local from RTDataColumns
        headerRow.createCell(75).setCellValue(DNColumns.TSTAMP_LOCAL.getColumnName()); // TSTAMP_LOCAL from DNColumns

        headerRow.createCell(76).setCellValue(RTDataColumns.PAN.getColumnName()); // acct_no from RTDataColumns
        headerRow.createCell(77).setCellValue(DNColumns.TSTAMP_LOCAL.getColumnName()); // TSTAMP_LOCAL from DNColumns

        headerRow.createCell(78).setCellValue(RTDataColumns.EXPIRY_DATE.getColumnName()); // expiry_dat from RTDataColumns
        headerRow.createCell(79).setCellValue(DNColumns.DATE_EXP.getColumnName()); // DATE_EXP from DNColumns

        headerRow.createCell(80).setCellValue(RTDataColumns.MSG_CLASS.getColumnName()); // merchant_type from RTDataColumns
        headerRow.createCell(81).setCellValue(DNColumns.MERCH_TYPE.getColumnName()); // MERCH_TYPE from DNColumns
    }

    // Processes the Realtime_Data sheet, matching transactions with DN_Data and writing results to the output sheet
    public static void processRealtimeSheet(Sheet rtSheet, Map<String, Row> dnDataMap, Sheet outputMappingSheet, Map<String, Integer> rtColumnMap, Map<String, Integer> dnColumnMap) {
        int outputRowNum = 1;
        for (Row row : rtSheet) {
            String tranNumberRealtime = getCellValueAsString(row.getCell(rtColumnMap.get(RTDataColumns.TRAN_NR.getColumnName())));
            if (dnDataMap.containsKey(tranNumberRealtime)) {
                Row outputRow = outputMappingSheet.createRow(outputRowNum++);
                Row dnDataRow = dnDataMap.get(tranNumberRealtime);

                // Populate the output row with relevant data from both sheets
                mapRowData(outputRow, dnDataRow, row, rtColumnMap, dnColumnMap);
            }
        }
    }

    // Maps the data from DN and Realtime rows into the output row using column names
    // Maps the data from DN and Realtime rows into the output row using column names
    private static void mapRowData(Row outputRow, Row dnDataRow, Row rtDataRow, Map<String, Integer> rtColumnMap, Map<String, Integer> dnColumnMap) {
        outputRow.createCell(0).setCellValue(getCellValueAsString(rtDataRow.getCell(rtColumnMap.get(RTDataColumns.MSG_TYPE.getColumnName())))); // msg_type from Realtime_Data
        outputRow.createCell(1).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get(DNColumns.MTI.getColumnName())))); // MTI from DN_Data
        outputRow.createCell(2).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get(DNColumns.FUNC_CODE.getColumnName())))); // FUNC_CODE from DN_Data
        outputRow.createCell(3).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get(DNColumns.TRAN_NUMBER.getColumnName())))); // TRAN_NUMBER from DN_Data
        outputRow.createCell(4).setCellValue(getCellValueAsString(rtDataRow.getCell(rtColumnMap.get(RTDataColumns.TRAN_NR.getColumnName())))); // tran_nr from Realtime_Data
        outputRow.createCell(5).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get(DNColumns.TRAN_DESC.getColumnName())))); // TRAN_DESC from DN_Data
        outputRow.createCell(6).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get(DNColumns.DRAFT_CAPTURE_FLG.getColumnName())))); // DRAFT_CAPTURE_FLG from DN_Data
        outputRow.createCell(7).setCellValue(getCellValueAsString(rtDataRow.getCell(rtColumnMap.get(RTDataColumns.DRAFT_CAPTURE.getColumnName())))); // DRAFT_CAPTURE from Realtime_Data
        outputRow.createCell(8).setCellValue(getCellValueAsString(rtDataRow.getCell(rtColumnMap.get(RTDataColumns.STAN_IN.getColumnName())))); // stan_in from Realtime_Data
        outputRow.createCell(9).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get(DNColumns.STANDIN_ACT.getColumnName())))); // STANDIN_ACT from DN_Data
        outputRow.createCell(10).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get(DNColumns.DATE_RECON_ACQ.getColumnName())))); // DATE_RECON_ACQ from DN_Data
        outputRow.createCell(11).setCellValue(getCellValueAsString(rtDataRow.getCell(rtColumnMap.get(RTDataColumns.SRCNODE_DATE_SETTLE.getColumnName())))); // srcnode_date_settle from Realtime_Data
        outputRow.createCell(12).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get(DNColumns.ADL_RQST_AMT1.getColumnName())))); // ADL_RQST_AMT1 from DN_Data
        outputRow.createCell(13).setCellValue(getCellValueAsString(rtDataRow.getCell(rtColumnMap.get(RTDataColumns.SRCNODE_CASH_REQUESTED.getColumnName())))); // srcnode_cash_requested from Realtime_Data
        outputRow.createCell(14).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get(DNColumns.CUR_RECON_ACQ_1.getColumnName())))); // CUR_RECON_ACQ_1 from DN_Data
        outputRow.createCell(15).setCellValue(getCellValueAsString(rtDataRow.getCell(rtColumnMap.get(RTDataColumns.SRCNODE_CURRENCY_CODE.getColumnName())))); // srcnode_currency_code from Realtime_Data
        outputRow.createCell(16).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get(DNColumns.CNV_RCN_ACQ_RATE.getColumnName())))); // CNV_RCN_ACQ_RATE from DN_Data
        outputRow.createCell(17).setCellValue(getCellValueAsString(rtDataRow.getCell(rtColumnMap.get(RTDataColumns.SRCNODE_CONVERSION_RATE.getColumnName())))); // srcnode_conversion_rate from Realtime_Data
        outputRow.createCell(18).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get(DNColumns.CNV_RCN_ACQ_DE_POS.getColumnName())))); // CNV_RCN_ACQ_DE_POS from DN_Data
        outputRow.createCell(19).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get(DNColumns.DATE_CNV_ACQ.getColumnName())))); // DATE_CNV_ACQ from DN_Data
        outputRow.createCell(20).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get(DNColumns.ODE_MTI.getColumnName())))); // ODE_MTI from DN_Data
        outputRow.createCell(21).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get(DNColumns.ODE_SYS_TRA_AUD_NO.getColumnName())))); // ODE_SYS_TRA_AUD_NO from DN_Data
        outputRow.createCell(22).setCellValue(getCellValueAsString(rtDataRow.getCell(rtColumnMap.get(RTDataColumns.SRCNODE_ORIGINAL_DATA.getColumnName())))); // srcnode_original_data from Realtime_Data
        outputRow.createCell(23).setCellValue(getCellValueAsString(rtDataRow.getCell(rtColumnMap.get(RTDataColumns.SNKNODE_REQ_SYS_TRACE.getColumnName())))); // snknode_req_sys_trace from Realtime_Data
        outputRow.createCell(24).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get(DNColumns.SYS_TRACE_AUDIT_NO.getColumnName())))); // SYS_TRACE_AUDIT_NO from DN_Data
        outputRow.createCell(25).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get(DNColumns.TRAN_DESC.getColumnName())))); // TRAN_DESC from DN_Data
        outputRow.createCell(26).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get(DNColumns.DATE_RECON_ISS.getColumnName())))); // DATE_RECON_ISS from DN_Data
        outputRow.createCell(27).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get(DNColumns.DATE_RECON_NET.getColumnName())))); // DATE_RECON_NET from DN_Data
        outputRow.createCell(28).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get(DNColumns.AMT_RECON_ISS.getColumnName())))); // AMT_RECON_ISS from DN_Data
        outputRow.createCell(29).setCellValue(getCellValueAsString(rtDataRow.getCell(rtColumnMap.get(RTDataColumns.SNKNODE_AMOUNT_REQUESTED.getColumnName())))); // snknode_amount_requested from Realtime_Data
        outputRow.createCell(30).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get(DNColumns.O_AMT_CARD_BILL.getColumnName())))); // O_AMT_CARD_BILL from DN_Data
        outputRow.createCell(31).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get(DNColumns.CUR_RECON_ISS.getColumnName())))); // CUR_RECON_ISS from DN_Data
        outputRow.createCell(33).setCellValue(getCellValueAsString(rtDataRow.getCell(rtColumnMap.get(RTDataColumns.SRCNODE_CASH_FINAL.getColumnName())))); // srcnode_cash_final from Realtime_Data
        outputRow.createCell(34).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get(DNColumns.CUR_RECON_NET.getColumnName())))); // CUR_RECON_NET from DN_Data
        outputRow.createCell(35).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get(DNColumns.CUR_CARD_BILL.getColumnName())))); // CUR_CARD_BILL from DN_Data
        outputRow.createCell(36).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get(DNColumns.CNV_RCN_ISS_RATE.getColumnName())))); // CNV_RCN_ISS_RATE from DN_Data
        outputRow.createCell(37).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get(DNColumns.DATE_CNV_ISS.getColumnName())))); // DATE_CNV_ISS from DN_Data
        outputRow.createCell(38).setCellValue(getCellValueAsString(dnDataRow.getCell(dnColumnMap.get(DNColumns.INST_ID_RECON_ISS.getColumnName())))); // INST_ID_RECON_ISS from DN_Data
        outputRow.createCell(39).setCellValue(getCellValueAsString(rtDataRow.getCell(rtColumnMap.get(RTDataColumns.TRAN_TYPE.getColumnName())))); // TRAN_TYPE from Realtime_Data
        outputRow.createCell(40).setCellValue(getCellValueAsString(rtDataRow.getCell(rtColumnMap.get(RTDataColumns.AMOUNT_TRAN_REQUESTED.getColumnName())))); // amount_tran_requested from Realtime_Data
        outputRow.createCell(41).setCellValue(getCellValueAsString(rtDataRow.getCell(rtColumnMap.get(RTDataColumns.AMOUNT_TRAN_APPROVED.getColumnName())))); // amount_tran_approved from Realtime_Data
        outputRow.createCell(42).setCellValue(getCellValueAsString(rtDataRow.getCell(rtColumnMap.get(RTDataColumns.GMT_DATE_TIME.getColumnName())))); // gmt_date_time from Realtime_Data
        outputRow.createCell(43).setCellValue(getCellValueAsString(rtDataRow.getCell(rtColumnMap.get(RTDataColumns.TIME_LOCAL.getColumnName())))); // time_local from Realtime_Data
        outputRow.createCell(44).setCellValue(getCellValueAsString(rtDataRow.getCell(rtColumnMap.get(RTDataColumns.DATE_LOCAL.getColumnName())))); // date_local from Realtime_Data
        outputRow.createCell(45).setCellValue(getCellValueAsString(rtDataRow.getCell(rtColumnMap.get(RTDataColumns.EXPIRY_DATE.getColumnName())))); // expiry_date from Realtime_Data
       // outputRow.createCell(46).setCellValue(getCellValueAsString(rtDataRow.getCell(rtColumnMap.get(RTDataColumns.MERCH_TYPE.getColumnName())))); // merch_type from Realtime_Data
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
