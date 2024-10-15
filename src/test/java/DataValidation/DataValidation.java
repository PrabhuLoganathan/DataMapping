package DataValidation;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;

public class DataValidation {

    public static void main(String[] args) throws IOException {

        DataValidatorUtils.initializeValidMtiMsgTypeMap();

        String inputFilePath = "/Users/prabhuloganathan/Downloads/DataMap/mappedDataSheet.xlsx";
        String outputFilePath = "/Users/prabhuloganathan/Downloads/DataMap/mappedDataSheet.xlsx"; // Use the same file

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

                    // Copy original row values
                   // DataValidatorUtils.copyRowValues(outputRow, row);


                    // Find the first empty cell in the row to avoid overlapping
                    //int lastCellIndex = row.getLastCellNum();
                    int lastCellIndex = 0;

                    // Fetch values by column name
                    String msgType = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get("msg_type")));
                    String mti = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get("MTI")));

                    // Validation 1: msg_type and MTI
                    boolean isValid = DataValidatorUtils.isValidMtiMsgType(msgType, mti);
                    outputRow.createCell(lastCellIndex).setCellValue(isValid ? "Valid" : "Invalid");
                    lastCellIndex++; // Move to next cell for the next validation

                    // Validation 2: Function Code
                    String funcCode = DataValidatorUtils.getFunctionCode(mti);
                    outputRow.createCell(lastCellIndex).setCellValue(funcCode);
                    lastCellIndex++;

                    // Validation 3: Draft Capture Flags
                    String draftCaptureFlgDn = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get("DRAFT_CAPTURE_FLG_DN")));
                    String draftCaptureFlgRt = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get("DRAFT_CAPTURE_FLG_RT")));
                    boolean isDraftCaptureValid = DataValidatorUtils.isValidDraftCaptureFlags(draftCaptureFlgDn, draftCaptureFlgRt);
                    outputRow.createCell(lastCellIndex).setCellValue(isDraftCaptureValid ? "Valid" : "Invalid");
                    lastCellIndex++;

                    // Validation 4: stan_in and STANDIN_ACT
                    String stanIn = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get("stan_in")));
                    String standinAct = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get("STANDIN_ACT")));
                    boolean isStanInStandinActValid = DataValidatorUtils.compareStanInAndStandinAct(stanIn, standinAct);
                    outputRow.createCell(lastCellIndex).setCellValue(isStanInStandinActValid ? "Valid" : "Invalid");
                    lastCellIndex++;

                    // Validation 5: DATE_RECON_ACQ and srcnode_date_settle
                    String dateReconAcq = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get("DATE_RECON_ACQ")));
                    String srcnodeDateSettle = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get("srcnode_date_settle")));
                    boolean isDateReconAcqSrcnodeDateSettleValid = DataValidatorUtils.compareDatesWithYearIgnored(dateReconAcq, srcnodeDateSettle);
                    outputRow.createCell(lastCellIndex).setCellValue(isDateReconAcqSrcnodeDateSettleValid ? "Valid" : "Invalid");
                    lastCellIndex++;

                    // Validation 6: ADL_RQST_AMT1 from DN Data and srcnode_cash_requested from Realtime Data
                    String additionalRequestAmount = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get("ADL_RQST_AMT1")));
                    String sourceNodeCashRequested = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get("srcnode_cash_requested")));
                    boolean isSourceNodeCashRequestedAndAdditionalRequestAmountValid = DataValidatorUtils.compareSourceNodeCashRequestedAndAdditionalRequestAmount(additionalRequestAmount, sourceNodeCashRequested);
                    outputRow.createCell(lastCellIndex).setCellValue(isSourceNodeCashRequestedAndAdditionalRequestAmountValid ? "Valid" : "Invalid");
                    lastCellIndex++;

                    // Validation 7: CUR_RECON_ACQ_1 from DN Data and srcnode_currency_code from Realtime Data
                    String curReconAcq1 = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get("CUR_RECON_ACQ_1")));
                    String srcnodeCurrencyCode = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get("srcnode_currency_code")));
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
