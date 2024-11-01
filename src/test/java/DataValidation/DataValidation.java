package DataValidation;

import DataMapping.DNColumns;
import DataMapping.RTDataColumns;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;

public class DataValidation {

    public static void main(String[] args) throws IOException {

        DataValidatorUtils.initializeValidMtiMsgTypeMap();

        String inputFilePath = "/Users/prabhuloganathan/Downloads/FinalData/mappedDataSheet.xlsx";
        String outputFilePath = "/Users/prabhuloganathan/Downloads/FinalData/mappedDataSheet.xlsx";

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
                    int lastCellIndex = 0;

                    // Validation 1: msg_type and MTI
                    String msgType = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.getOrDefault(RTDataColumns.MSG_TYPE.getColumnName(), -1)));
                    String mti = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.getOrDefault(DNColumns.MTI.getColumnName(), -1)));
                    boolean isValid = DataValidatorUtils.isValidMtiMsgType(msgType, mti);
                    outputRow.createCell(lastCellIndex++).setCellValue(isValid ? "Valid" : "Invalid");

                    // Validation 2: Function Code
                    String funcCode = DataValidatorUtils.getFunctionCode(mti);
                    outputRow.createCell(lastCellIndex++).setCellValue(funcCode);

                    // Validation 3: Draft Capture Flags
                    String draftCaptureFlgDn = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.getOrDefault(DNColumns.DRAFT_CAPTURE_FLG.getColumnName(), -1)));
                    String draftCaptureFlgRt = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.getOrDefault(RTDataColumns.DRAFT_CAPTURE.getColumnName(), -1)));
                    boolean isDraftCaptureValid = DataValidatorUtils.isValidDraftCaptureFlags(draftCaptureFlgDn, draftCaptureFlgRt);
                    outputRow.createCell(lastCellIndex++).setCellValue(isDraftCaptureValid ? "Valid" : "Invalid");

                    // Validation 4: stan_in and STANDIN_ACT
                    String stanIn = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.getOrDefault(RTDataColumns.STAN_IN.getColumnName(), -1)));
                    String standinAct = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.getOrDefault(DNColumns.STANDIN_ACT.getColumnName(), -1)));
                    boolean isStanInStandinActValid = DataValidatorUtils.compareStanInAndStandinAct(stanIn, standinAct);
                    outputRow.createCell(lastCellIndex++).setCellValue(isStanInStandinActValid ? "Valid" : "Invalid");

                    // Validation 5: DATE_RECON_ACQ and srcnode_date_settle
                    String dateReconAcq = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.getOrDefault(DNColumns.DATE_RECON_ACQ.getColumnName(), -1)));
                    String srcnodeDateSettle = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.getOrDefault(RTDataColumns.SRCNODE_DATE_SETTLE.getColumnName(), -1)));
                    boolean isDateReconAcqSrcnodeDateSettleValid = DataValidatorUtils.compareDatesWithYearIgnored(dateReconAcq, srcnodeDateSettle);
                    outputRow.createCell(lastCellIndex++).setCellValue(isDateReconAcqSrcnodeDateSettleValid ? "Valid" : "Invalid");

                    // Validation 6: ADL_RQST_AMT1 and srcnode_cash_requested
                    String additionalRequestAmount = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.getOrDefault(DNColumns.ADL_RQST_AMT1.getColumnName(), -1)));
                    String sourceNodeCashRequested = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.getOrDefault(RTDataColumns.SRCNODE_CASH_REQUESTED.getColumnName(), -1)));
                    boolean isSourceNodeCashRequestedAndAdditionalRequestAmountValid = DataValidatorUtils.compareSourceNodeCashRequestedAndAdditionalRequestAmount(additionalRequestAmount, sourceNodeCashRequested);
                    outputRow.createCell(lastCellIndex++).setCellValue(isSourceNodeCashRequestedAndAdditionalRequestAmountValid ? "Valid" : "Invalid");

                    // Validation 7: CUR_RECON_ACQ_1 and srcnode_currency_code
                    String curReconAcq1 = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.getOrDefault(DNColumns.CUR_RECON_ACQ_1.getColumnName(), -1)));
                    String srcnodeCurrencyCode = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.getOrDefault(RTDataColumns.SRCNODE_CURRENCY_CODE.getColumnName(), -1)));
                    boolean isCurReconAcq1AndSrcnodeCurrencyCodeValid = DataValidatorUtils.compareCurReconAcqAndSrcnodeCurrencyCode(curReconAcq1, srcnodeCurrencyCode);
                    outputRow.createCell(lastCellIndex++).setCellValue(isCurReconAcq1AndSrcnodeCurrencyCodeValid ? "Valid" : "Invalid");

                    // Validation 8: Conversion Rates
                    String srcnodeConversionRate = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.getOrDefault("srcnode_conversion_rate", -1)));
                    String cnvRcnAcqDePos = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.getOrDefault("CNV_RCN_ACQ_DE_POS", -1)));
                    String cnvRcnAcqRate = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.getOrDefault("CNV_RCN_ACQ_RATE", -1)));
                    boolean isConversionRateValid = DataValidatorUtils.validateConversionRates(srcnodeConversionRate, cnvRcnAcqDePos, cnvRcnAcqRate);
                    outputRow.createCell(lastCellIndex++).setCellValue(isConversionRateValid ? "Valid" : "Invalid");

                    // Validation 9: DATE_CNV_ACQ and srcnode_date_conversion
                    String dateCnvAcq = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.getOrDefault(DNColumns.DATE_CNV_ACQ.getColumnName(), -1)));
                    String srcNodeDateConversion = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.getOrDefault(RTDataColumns.SRCNODE_DATE_CONVERSION.getColumnName(), -1)));
                    boolean isDateCnvAcqAndSrcNodeDateConversion = DataValidatorUtils.compareDatesWithYearIgnored(dateCnvAcq, srcNodeDateConversion);
                    outputRow.createCell(lastCellIndex++).setCellValue(isDateCnvAcqAndSrcNodeDateConversion ? "Valid" : "Invalid");


                    // Validation 10: O_AMT_RECON_ACQ and snknode_amount_requested for MTI 1430
                    String oAmtReconAcq = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get("O_AMT_RECON_ACQ")));
                    String snkNodeAmountRequested = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get("snknode_amount_requested")));
                    boolean isAmtReconAcqValid = DataValidatorUtils.validateAmtReconAcqForMti(mti, oAmtReconAcq, snkNodeAmountRequested);
                    outputRow.createCell(lastCellIndex++).setCellValue(isAmtReconAcqValid ? "Valid" : "Invalid");


                    // Validation 11: snknode_rev_sys_trace and SYS_TRACE_AUDIT_NO for MTI 1430
//                    String MTI = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get(DNColumns.MTI.getColumnName())));
                    String snknodeRevSysTrace = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get("snknode_rev_sys_trace")));
                    String sysTraceAuditNo = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get(DNColumns.SYS_TRACE_AUDIT_NO.getColumnName())));
                    boolean isSysTraceAuditValid = DataValidatorUtils.validateSysTraceAuditForMti(mti, snknodeRevSysTrace, sysTraceAuditNo);
                    outputRow.createCell(lastCellIndex++).setCellValue(isSysTraceAuditValid ? "Valid" : "Invalid");

                    // Validation 12: snknode_req_sys_trace and SYS_TRACE_AUDIT_NO for MTI 1110, 1210, 1410
                    String snknodeReqSysTrace = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get("snknode_req_sys_trace")));
                    boolean isReqSysTraceAuditValid = DataValidatorUtils.validateReqSysTraceForMti(mti, snknodeReqSysTrace, sysTraceAuditNo);
                    outputRow.createCell(lastCellIndex++).setCellValue(isReqSysTraceAuditValid ? "Valid" : "Invalid");

                    // Validation 13: snknode_adv_sys_trace and SYS_TRACE_AUDIT_NO for MTI 1230
                    String snknodeAdvSysTrace = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get("snknode_adv_sys_trace")));
                    boolean isAdvSysTraceAuditValid = DataValidatorUtils.validateAdvSysTraceForMti(mti, snknodeAdvSysTrace, sysTraceAuditNo);
                    outputRow.createCell(lastCellIndex++).setCellValue(isAdvSysTraceAuditValid ? "Valid" : "Invalid");

                    // Validation 14: O_AMT_CARD_BILL, O_AMT_RECON_NET, O_AMT_RECON_ISS should match snknode_amount_requested for MTI 1430

                    String oAmtCardBill = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get("O_AMT_CARD_BILL")));
                    String oAmtReconNet = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get("O_AMT_RECON_NET")));
                    // Coloumn  name is not avialble
                    String oAmtReconIss = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get("O_AMT_RECON_NET")));
                    String snknodeAmountRequested = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get("snknode_amount_requested")));

                    boolean isValidation14Passed = DataValidatorUtils.validateAmountsForMti1430(mti, oAmtCardBill, oAmtReconNet, oAmtReconIss, snknodeAmountRequested);
                    outputRow.createCell(lastCellIndex++).setCellValue(isValidation14Passed ? "Valid" : "Invalid");


                    // Validaion 15
                    // sikipped


                    // Validation 16: CUR_RECON_ISS, CUR_RECON_NET, CUR_CARD_BILL should match snknode_currency_code
                    String curReconIss = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get(DNColumns.CUR_RECON_ISS.getColumnName())));
                    String curReconNet = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get(DNColumns.CUR_RECON_NET.getColumnName())));
                    String curCardBill = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get(DNColumns.CUR_CARD_BILL.getColumnName())));
                    String snknodeCurrencyCode = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get(RTDataColumns.SNKNODE_CURRENCY_CODE.getColumnName())));

                    boolean isValidation16Passed = DataValidatorUtils.validateCurrencies(curReconIss, curReconNet, curCardBill, snknodeCurrencyCode);
                    outputRow.createCell(lastCellIndex++).setCellValue(isValidation16Passed ? "Valid" : "Invalid");

                    // Validation 17: snknode_date_conversion and DATE_CNV_ISS should be equal
                    String snknodeDateConversion = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get(RTDataColumns.SNKNODE_DATE_CONVERSION.getColumnName())));
                    String dateCnvIss = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get(DNColumns.DATE_CNV_ISS.getColumnName())));

                    boolean isValidation17Passed = DataValidatorUtils.validateDateConversion(snknodeDateConversion, dateCnvIss);
                    outputRow.createCell(lastCellIndex++).setCellValue(isValidation17Passed ? "Valid" : "Invalid");

//
//                    // Validation 18: Mapping validation using modified totals_group and currency code "AU"
//                    String totalsGroup = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get("totals_group")));
//                    String instIdReconIss = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get(DNColumns.INST_ID_RECON_ISS.getColumnName())));
//                    String instIdRecnIssB = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get(DNColumns.INST_ID_RECN_ISS_B.getColumnName())));
//
//                    boolean isValidation18Passed = DataValidatorUtils.validateInstIdMapping(totalsGroup, "AU", instIdReconIss, instIdRecnIssB);
//                    outputRow.createCell(lastCellIndex++).setCellValue(isValidation18Passed ? "Valid" : "Invalid");

                    // Validation 22: amount_tran_requested, AMT_TRAN, and O_AMT_TRAN should match if MTI is 1430
                    String amountTranRequested = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get(RTDataColumns.AMOUNT_TRAN_REQUESTED.getColumnName())));
                    String amtTran = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get(DNColumns.AMT_TRAN.getColumnName())));
                    String oAmtTran = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get(DNColumns.O_AMT_TRAN.getColumnName())));

                    boolean isValidation22Passed = DataValidatorUtils.validateAmountsForMti1430(mti, amountTranRequested, amtTran, oAmtTran);
                    outputRow.createCell(lastCellIndex++).setCellValue(isValidation22Passed ? "Valid" : "Invalid");

                    // Validation 24: gmt_date_time (MMDDHHMMSS) and GMT_time (YYYYMMDDHHMMSS) should match
                    String gmtDateTime = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get(RTDataColumns.GMT_DATE_TIME.getColumnName())));
                    String gmtTime = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get(DNColumns.GMT_TIME.getColumnName())));

                    boolean isValidation24Passed = DataValidatorUtils.validateGmtDateTime(gmtDateTime, gmtTime);
                    outputRow.createCell(lastCellIndex++).setCellValue(isValidation24Passed ? "Valid" : "Invalid");

                    // Validation 25: time_local (HHMMSS) and TSTAMP_LOCAL (YYYYMMDDHHMMSS) should match
                    String timeLocal = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get(RTDataColumns.TIME_LOCAL.getColumnName())));
                    String tstampLocal = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get(DNColumns.TSTAMP_LOCAL.getColumnName())));


                    boolean isValidation25Passed = DataValidatorUtils.validateTimeLocal(timeLocal, tstampLocal);
                    outputRow.createCell(lastCellIndex++).setCellValue(isValidation25Passed ? "Valid" : "Invalid");


                    // Validation 26: date_local (MMDD) and TSTAMP_LOCAL (YYYYMMDDHHMMSS) should match
                    String dateLocal = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get(RTDataColumns.DATE_LOCAL.getColumnName())));
                    String tstampLocal2 = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get(DNColumns.TSTAMP_LOCAL.getColumnName())));

                    boolean isValidation26Passed = DataValidatorUtils.validateDateLocal(dateLocal, tstampLocal2);


                    // Validation 28: merchant_type and MERCH_TYPE should have the same value
                    String merchantType = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get("MERCH_TYPE")));
                    String merchType = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get(DNColumns.MERCH_TYPE.getColumnName())));

                    boolean isValidation28Passed = DataValidatorUtils.validateMerchantType(merchantType, merchType);
                    outputRow.createCell(lastCellIndex++).setCellValue(isValidation28Passed ? "Valid" : "Invalid");

//                    // Validation 30: card_acceptor_term_id, CARD_ACPT_TERM_ID, and NET_TERM_ID should match
//                    String cardAcceptorTermId = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get("CARD_ACCEPTOR_TERM_ID")));
//                    String cardAcptTermId = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get("CARD_ACPT_TERM_ID")));
//                    String netTermId = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get("NET_TERM_ID")));
//
//                    boolean isValidation30Passed = DataValidatorUtils.validateTerminalIds(cardAcceptorTermId, cardAcptTermId, netTermId);
//                    outputRow.createCell(lastCellIndex++).setCellValue(isValidation30Passed ? "Valid" : "Invalid");


//                    // Validation 31: Card_acceptor_id_code and CARD_ACPT_ID should match
//                    String cardAcceptorIdCode = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get("CARD_ACCEPTOR_ID_CODE")));
//                    String cardAcptId = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get("CARD_ACPT_ID")));
//
//                    boolean isValidation31Passed = DataValidatorUtils.validateCardAcceptorId(cardAcceptorIdCode, cardAcptId);
//                    outputRow.createCell(lastCellIndex++).setCellValue(isValidation31Passed ? "Valid" : "Invalid");

//                    // Validation 32: Match card_acceptor_name_loc with CARD_ACPT_NAME_LOC
//                    String cardAcceptorNameLoc = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get("card_acceptor_name_loc")));
//                    String cardAcptNameLoc = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get("CARD_ACPT_NAME_LOC")));
//                    boolean isNameLocValid = DataValidatorUtils.validateNameLoc(cardAcceptorNameLoc, cardAcptNameLoc);
//                    outputRow.createCell(lastCellIndex++).setCellValue(isNameLocValid ? "Valid" : "Invalid");

                    // Validation 32: Country Code Check for CARD_ACPT_COUNTRY
                    String cardAcptCountry = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get("CARD_ACPT_COUNTRY")));
                    boolean isCountryCodeValid = DataValidatorUtils.validateCountryCode(cardAcptCountry, "AUS");
                    outputRow.createCell(lastCellIndex++).setCellValue(isCountryCodeValid ? "Valid" : "Invalid");

                    // Validation 32: Composite Code Check for card_acceptor_id_code format
                    String cardAcceptorIdCode2 = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get("card_acceptor_id_code")));
                    boolean isIdCodeFormatValid = DataValidatorUtils.validateIdCodeFormat(cardAcceptorIdCode2, "02AU");
                    outputRow.createCell(lastCellIndex++).setCellValue(isIdCodeFormatValid ? "Valid" : "Invalid");


//                    // Validation 33: Validate each position in pos_data_code
//                    String posDataCode = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get("pos_data_code")));
//                    boolean isValidation33Passed = DataValidatorUtils.validatePosDataCodeMapping(posDataCode, row, columnMap);
//                    outputRow.createCell(lastCellIndex++).setCellValue(isValidation33Passed ? "Valid" : "Invalid");

                    // Validation 39: Snknode_acquiring_inst and INST_ID_ISS are the same
//                    String snknodeAcquiringInst = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get("Snknode_acquiring_inst")));
//                    String instIdIss = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get("INST_ID_ISS")));
//                    boolean isSnknodeAcquiringInstAndInstIdIssValid = DataValidatorUtils.validateSnknodeAcquiringInstAndInstIdIss(snknodeAcquiringInst, instIdIss);
//                    outputRow.createCell(lastCellIndex++).setCellValue(isSnknodeAcquiringInstAndInstIdIssValid ? "Valid" : "Invalid");

//                    // Validation 40: Card_verification_result is null and CVV_CVN_RESULT is empty
//                    String cardVerificationResult = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get("Card_verification_result")));
//                    String cvvCvnResult = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get("CVV_CVN_RESULT")));
//                    boolean isCardVerificationAndCVVValid = DataValidatorUtils.validateCardVerificationAndCVVResult(cardVerificationResult, cvvCvnResult);
//                    outputRow.createCell(lastCellIndex++).setCellValue(isCardVerificationAndCVVValid ? "Valid" : "Invalid");

//                    // Validation 41: Secure_3d_result is null and CAVV_RESULT is empty
//                    String secure3dResult = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get("Secure_3d_result")));
//                    String cavvResult = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get("CAVV_RESULT")));
//                    boolean isSecure3DAndCAVVValid = DataValidatorUtils.validateSecure3DAndCAVVResult(secure3dResult, cavvResult);
//                    outputRow.createCell(lastCellIndex++).setCellValue(isSecure3DAndCAVVValid ? "Valid" : "Invalid");


//                    // Validation 42: Decrypt pan_encrypted and compare with PAN
//                    String panEncrypted = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get("pan_encrypted")));
//                    String pan = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get("PAN")));
//                    boolean isPanValid = DataValidatorUtils.validateDecryptedPan(panEncrypted, pan);
//                    outputRow.createCell(lastCellIndex++).setCellValue(isPanValid ? "Valid" : "Invalid");

                    // Validation 43: Check if ret_ref_no and RETRIEVAL_REF_NO are identical
                    String retRefNo = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get("ret_ref_no")));
                    String retrievalRefNo = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get("RETRIEVAL_REF_NO")));
                    boolean isValidation43Passed = DataValidatorUtils.validateRetRefNo(retRefNo, retrievalRefNo);
                    outputRow.createCell(lastCellIndex++).setCellValue(isValidation43Passed ? "Valid" : "Invalid");


                    // Validation 45: Check if Msg_reason_code_rev and MSG_RESON_CODE_ACQ are equal when MTI is 1430
//
//                    String msgReasonCodeRev = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get("Msg_reason_code_rev")));
//                    String msgResonCodeAcq = DataValidatorUtils.getCellValueAsString(row.getCell(columnMap.get("MSG_RESON_CODE_ACQ")));
//
//                    boolean isValidation45Passed = DataValidatorUtils.validateMsgReasonCodeForMti1430(mti, msgReasonCodeRev, msgResonCodeAcq);
//                    outputRow.createCell(lastCellIndex++).setCellValue(isValidation45Passed ? "Valid" : "Invalid");

                    // Continue adding other validations (10 - 43) here
                    // Use `getOrDefault(columnName, -1)` for each column to avoid NullPointerExceptions
                    // Increment `lastCellIndex` for each cell created in `outputRow`

                    // Final output to workbook
                }
                rowNum++;
            }

            try (FileOutputStream fos = new FileOutputStream(outputFilePath)) {
                workbook.write(fos);
            }

            System.out.println("Validation completed. Output saved to: " + outputFilePath);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
