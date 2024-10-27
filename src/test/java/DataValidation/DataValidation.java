package DataValidation;

import DataMapping.DNColumns;
import DataMapping.RTDataColumns;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;

public class DataValidation {

    public static void main(String[] args) throws IOException {

        DataValidatorUtils.initializeValidMtiMsgTypeMap();

        String inputFilePath = "/Users/prabhuloganathan/IdeaProjects/DataMapping/src/test/testData/mappedDataSheet.xlsx";
        String outputFilePath = "/Users/prabhuloganathan/IdeaProjects/DataMapping/src/test/testData/mappedDataSheet.xlsx"; // Use the same file

        try (FileInputStream fis = new FileInputStream(inputFilePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            // Check if "ValidationResult" sheet exists, remove it if found
            int sheetIndex = workbook.getSheetIndex("ValidationResult");
            if (sheetIndex != -1) {
                workbook.removeSheetAt(sheetIndex);
            }

            Sheet sheet = workbook.getSheetAt(0);  // Read the original sheet
            Sheet outputSheet = workbook.createSheet("ValidationResult");  // Create a new sheet for validation

            // Get the header row to map column names to indexes
            Row headerRow = sheet.getRow(0);
            Map<String, Integer> columnMap = DataValidatorUtils.getColumnMap(headerRow);

            int rowNum = 0;
            for (Row row : sheet) {
                if (rowNum == 0) {
                    DataValidatorUtils.copyHeaderRow(outputSheet, row); // Copy header row and rename columns
                } else {
                    Row outputRow = outputSheet.createRow(rowNum);  // Create corresponding row in the output sheet


                     //DataValidatorUtils.copyRowValues(outputRow, row);

                    int lastCellIndex = 0;

                    // Validation 1: msg_type and MTI
                    String msgType = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get(RTDataColumns.MSG_TYPE.getColumnName())));
                    String mti = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get(DNColumns.MTI.getColumnName())));

                    boolean isValid = DataValidatorUtils.isValidMtiMsgType(msgType, mti);
                    outputRow.createCell(lastCellIndex).setCellValue(isValid ? "Valid" : "Invalid");
                    lastCellIndex++;

                    // Validation 2: Function Code
                    String funcCode = DataValidatorUtils.getFunctionCode(mti);
                    outputRow.createCell(lastCellIndex).setCellValue(funcCode);
                    lastCellIndex++;

                    // Validation 3: Draft Capture Flags
                    String draftCaptureFlgDn = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get(DNColumns.DRAFT_CAPTURE_FLG.getColumnName() + "_DN")));
                    String draftCaptureFlgRt = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get(RTDataColumns.DRAFT_CAPTURE.getColumnName())));
                    boolean isDraftCaptureValid = DataValidatorUtils.isValidDraftCaptureFlags(draftCaptureFlgDn, draftCaptureFlgRt);
                    outputRow.createCell(lastCellIndex).setCellValue(isDraftCaptureValid ? "Valid" : "Invalid");
                    lastCellIndex++;

                    // Validation 4: stan_in and STANDIN_ACT
                    String stanIn = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get(RTDataColumns.STAN_IN.getColumnName())));
                    String standinAct = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get(DNColumns.STANDIN_ACT.getColumnName())));
                    boolean isStanInStandinActValid = DataValidatorUtils.compareStanInAndStandinAct(stanIn, standinAct);
                    outputRow.createCell(lastCellIndex).setCellValue(isStanInStandinActValid ? "Valid" : "Invalid");
                    lastCellIndex++;

                    // Validation 5: DATE_RECON_ACQ and srcnode_date_settle
                    String dateReconAcq = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get(DNColumns.DATE_RECON_ACQ.getColumnName())));
                    String srcnodeDateSettle = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get(RTDataColumns.SRCNODE_DATE_SETTLE.getColumnName())));
                    boolean isDateReconAcqSrcnodeDateSettleValid = DataValidatorUtils.compareDatesWithYearIgnored(dateReconAcq, srcnodeDateSettle);
                    outputRow.createCell(lastCellIndex).setCellValue(isDateReconAcqSrcnodeDateSettleValid ? "Valid" : "Invalid");
                    lastCellIndex++;

                    // Validation 6: ADL_RQST_AMT1 from DN Data and srcnode_cash_requested from Realtime Data
                    String additionalRequestAmount = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get(DNColumns.ADL_RQST_AMT1.getColumnName())));
                    String sourceNodeCashRequested = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get(RTDataColumns.SRCNODE_CASH_REQUESTED.getColumnName())));
                    boolean isSourceNodeCashRequestedAndAdditionalRequestAmountValid = DataValidatorUtils.compareSourceNodeCashRequestedAndAdditionalRequestAmount(additionalRequestAmount, sourceNodeCashRequested);
                    outputRow.createCell(lastCellIndex).setCellValue(isSourceNodeCashRequestedAndAdditionalRequestAmountValid ? "Valid" : "Invalid");
                    lastCellIndex++;

                    // Validation 7: CUR_RECON_ACQ_1 from DN Data and srcnode_currency_code from Realtime Data
                    String curReconAcq1 = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get(DNColumns.CUR_RECON_ACQ_1.getColumnName())));
                    String srcnodeCurrencyCode = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get(RTDataColumns.SRCNODE_CURRENCY_CODE.getColumnName())));
                    boolean isCurReconAcq1AndSrcnodeCurrencyCodeValid = DataValidatorUtils.compareCurReconAcqAndSrcnodeCurrencyCode(curReconAcq1, srcnodeCurrencyCode);
                    outputRow.createCell(lastCellIndex).setCellValue(isCurReconAcq1AndSrcnodeCurrencyCodeValid ? "Valid" : "Invalid");
                }
                rowNum++;
            }

            try (FileOutputStream fos = new FileOutputStream(outputFilePath)) {
                workbook.write(fos); // Write the modified workbook back to the same file
            }

            System.out.println("Validation completed. Output saved to: " + outputFilePath);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}