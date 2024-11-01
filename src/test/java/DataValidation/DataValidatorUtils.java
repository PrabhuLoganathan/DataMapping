package DataValidation;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import javax.crypto.Cipher;
import javax.crypto.spec.SecretKeySpec;
import java.util.*;

public class DataValidatorUtils {

    private static final Map<String, List<Integer>> validMtiMsgTypeMap = new HashMap<>();
    private static final Map<String, Integer> funcCodeMap = new HashMap<>();

    public static void initializeValidMtiMsgTypeMap() {
        validMtiMsgTypeMap.put("1430", Arrays.asList(1057, 1056));
        validMtiMsgTypeMap.put("1110", List.of(256));
        validMtiMsgTypeMap.put("1210", List.of(512));
        validMtiMsgTypeMap.put("1230", Arrays.asList(544, 545));

        funcCodeMap.put("1210", 200);
        funcCodeMap.put("1230", 200);
        funcCodeMap.put("1110", 100);
        funcCodeMap.put("1430", 400);
        funcCodeMap.put("1410", 400);
    }

    public static Map<String, Integer> getColumnMap(Row headerRow) {
        Map<String, Integer> columnMap = new HashMap<>();
        for (Cell cell : headerRow) {
            columnMap.put(cell.getStringCellValue(), cell.getColumnIndex());
        }
        return columnMap;
    }

    public static void copyRowValues(Row outputRow, Row inputRow) {
        for (int i = 0; i < inputRow.getPhysicalNumberOfCells(); i++) {
            Cell inputCell = inputRow.getCell(i);
            Cell outputCell = outputRow.createCell(i);
            if (inputCell != null) {
                switch (inputCell.getCellType()) {
                    case STRING -> outputCell.setCellValue(inputCell.getStringCellValue());
                    case NUMERIC -> outputCell.setCellValue(inputCell.getNumericCellValue());
                    case BOOLEAN -> outputCell.setCellValue(inputCell.getBooleanCellValue());
                    case FORMULA -> outputCell.setCellFormula(inputCell.getCellFormula());
                    default -> outputCell.setCellValue("");
                }
            }
        }
    }

    public static void copyHeaderRow(Sheet outputSheet, Row headerRow) {
        Row outputHeaderRow = outputSheet.createRow(0);
        for (int i = 0; i < headerRow.getPhysicalNumberOfCells(); i++) {
            Cell inputCell = headerRow.getCell(i);
            Cell outputCell = outputHeaderRow.createCell(i);
            if (inputCell != null) {
                //outputCell.setCellValue(inputCell.getStringCellValue());
            }
        }

//        int colIndex = headerRow.getPhysicalNumberOfCells();
        int colIndex =0;

        outputHeaderRow.createCell(colIndex++).setCellValue("msg_type and MTI validation");
        outputHeaderRow.createCell(colIndex++).setCellValue("Function Code");
        outputHeaderRow.createCell(colIndex++).setCellValue("Draft Capture Validation");
        outputHeaderRow.createCell(colIndex++).setCellValue("Validate_StanIn_StandinAct");
        outputHeaderRow.createCell(colIndex++).setCellValue("Validate_DATE_RECON_ACQ_srcnode_date_settle");
        outputHeaderRow.createCell(colIndex++).setCellValue("Validate_ADL_RQST_AMT1_srcnode_cash_requested");
        outputHeaderRow.createCell(colIndex++).setCellValue("Validate_CUR_RECON_ACQ_1_srcnode_currency_code");
        outputHeaderRow.createCell(colIndex++).setCellValue("Validate_Conversion_Rates");
        outputHeaderRow.createCell(colIndex).setCellValue("Validate_O_AMT_RECON_ACQ_snknode_amount_requested");
    }

    public static boolean isValidMtiMsgType(String msgType, String mti) {
        if (msgType == null || msgType.isEmpty() || mti == null || mti.isEmpty()) {
            return false;
        }
        try {
            int msgTypeValue = Integer.parseInt(msgType);
            return validMtiMsgTypeMap.getOrDefault(mti, List.of()).contains(msgTypeValue);
        } catch (NumberFormatException e) {
            System.out.println("Error parsing msgType or MTI: " + msgType + ", " + mti);
        }
        return false;
    }

    public static boolean isValidDraftCaptureFlags(String draftFlgDn, String draftFlgRt) {
        return (draftFlgDn.equals("0") || draftFlgDn.equals("1")) && (draftFlgRt.equals("0") || draftFlgRt.equals("1"));
    }

    public static String getFunctionCode(String mti) {
        return funcCodeMap.containsKey(mti) ? String.valueOf(funcCodeMap.get(mti)) : "Not Found";
    }

    public static boolean compareStanInAndStandinAct(String stanIn, String standinAct) {
        return stanIn.equals(standinAct);
    }

    public static boolean compareSourceNodeCashRequestedAndAdditionalRequestAmount(String additionalRequestAmount, String sourceNodeCashRequested) {
        return additionalRequestAmount.equals(sourceNodeCashRequested);
    }

    public static boolean compareCurReconAcqAndSrcnodeCurrencyCode(String curReconAcq1, String srcnodeCurrencyCode) {
        return curReconAcq1.equals(srcnodeCurrencyCode);
    }

    public static boolean compareDatesWithYearIgnored(String dateReconAcq, String srcnodeDateSettle) {
        if (dateReconAcq.length() >= 8 && srcnodeDateSettle.length() == 3) {
            String dateReconAcqMMDD = dateReconAcq.substring(5);
            return dateReconAcqMMDD.equals(srcnodeDateSettle);
        }
        return false;
    }

    public static String getCellValueAsString(Cell cell) {
        if (cell == null) return "";
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> String.valueOf((int) cell.getNumericCellValue());
            default -> "";
        };
    }

    public static boolean validateConversionRates(String srcnodeConversionRate, String cnvRcnAcqDePos, String cnvRcnAcqRate) {
        try {
            double srcnodeRate = Double.parseDouble(srcnodeConversionRate);
            double acqRate = Double.parseDouble(cnvRcnAcqRate);
            int decimalPos = Integer.parseInt(cnvRcnAcqDePos);
            double adjustedAcqRate = acqRate / Math.pow(10, decimalPos);
            return srcnodeRate == adjustedAcqRate;
        } catch (NumberFormatException e) {
            System.out.println("Error parsing conversion rate data: " + e.getMessage());
            return false;
        }
    }

    // Validation 10 helper: Check if MTI is 1430 and O_AMT_RECON_ACQ matches snknode_amount_requested
    public static boolean validateAmtReconAcqForMti(String mti, String oAmtReconAcq, String snkNodeAmountRequested) {
        return "1430".equals(mti) && oAmtReconAcq.equals(snkNodeAmountRequested);
    }

    // Validation 11 helper: Check if MTI is 1430 and snknode_rev_sys_trace matches SYS_TRACE_AUDIT_NO
    public static boolean validateSysTraceAuditForMti(String mti, String snknodeRevSysTrace, String sysTraceAuditNo) {
        return "1430".equals(mti) && snknodeRevSysTrace.equals(sysTraceAuditNo);
    }

    // Validation 12 helper: Check if MTI is 1110, 1210, or 1410 and snknode_req_sys_trace matches SYS_TRACE_AUDIT_NO
    public static boolean validateReqSysTraceForMti(String mti, String snknodeReqSysTrace, String sysTraceAuditNo) {
        return ("1110".equals(mti) || "1210".equals(mti) || "1410".equals(mti)) && snknodeReqSysTrace.equals(sysTraceAuditNo);
    }

    // Validation 13 helper: Check if MTI is 1230 and snknode_adv_sys_trace matches SYS_TRACE_AUDIT_NO
    public static boolean validateAdvSysTraceForMti(String mti, String snknodeAdvSysTrace, String sysTraceAuditNo) {
        return "1230".equals(mti) && snknodeAdvSysTrace.equals(sysTraceAuditNo);
    }


    // Validation 14 helper: Check if MTI is 1430 and O_AMT_CARD_BILL, O_AMT_RECON_NET, O_AMT_RECON_ISS are equal to snknode_amount_requested
    public static boolean validateAmountsForMti1430(String mti, String oAmtCardBill, String oAmtReconNet, String oAmtReconIss, String snknodeAmountRequested) {
        return "1430".equals(mti)
                && oAmtCardBill.equals(oAmtReconNet)
                && oAmtReconNet.equals(oAmtReconIss)
                && oAmtReconIss.equals(snknodeAmountRequested);
    }


    // Validation 16 helper: Check if CUR_RECON_ISS, CUR_RECON_NET, CUR_CARD_BILL are equal and match snknode_currency_code
    public static boolean validateCurrencies(String curReconIss, String curReconNet, String curCardBill, String snknodeCurrencyCode) {
        return curReconIss.equals(curReconNet) && curReconNet.equals(curCardBill) && curCardBill.equals(snknodeCurrencyCode);
    }


    // Validation 17 helper: Check if snknode_date_conversion and DATE_CNV_ISS are equal
    public static boolean validateDateConversion(String snknodeDateConversion, String dateCnvIss) {
        return snknodeDateConversion.equals(dateCnvIss);
    }


    private static final Map<String, String[]> instIdMappingTable = new HashMap<>();

    static {
        // Populate mapping table with values from the image
        instIdMappingTable.put("AMEX+AU", new String[]{"AMEXAU", "AMEXAU"});
        instIdMappingTable.put("DINERS+AU", new String[]{"DINERSAU", "DINERSAU"});
        instIdMappingTable.put("ANZ+AU", new String[]{"ANZAU", "ANZAU"});
    }

    // Validation 18 helper: Check if derived INST_ID_RECON_ISS and INST_ID_RECN_ISS_B are valid
    public static boolean validateInstIdMapping(String totalsGroup, String currencyCode, String instIdReconIss, String instIdRecnIssB) {
        // Remove the first character from totalsGroup
        String modifiedTotalsGroup = totalsGroup.length() > 1 ? totalsGroup.substring(1) : totalsGroup;
        String key = modifiedTotalsGroup + "+" + currencyCode;

        if (instIdMappingTable.containsKey(key)) {
            String[] expectedInstIds = instIdMappingTable.get(key);
            return expectedInstIds[0].equals(instIdReconIss) && expectedInstIds[1].equals(instIdRecnIssB);
        }
        return false; // Return false if mapping is not found
    }

    // Validation 22 helper: Check if O_AMT_TRAN matches amount_tran_requested and AMT_TRAN when MTI is 1430
    public static boolean validateAmountsForMti1430(String mti, String amountTranRequested, String amtTran, String oAmtTran) {
        if (!"1430".equals(mti)) {
            return true; // Skip validation if MTI is not 1430
        }
        return oAmtTran.equals(amountTranRequested) && oAmtTran.equals(amtTran);
    }


    // Validation 24 helper: Check if gmt_date_time (MMDDHHMMSS) matches GMT_time (YYYYMMDDHHMMSS)
    public static boolean validateGmtDateTime(String gmtDateTime, String gmtTime) {
        if (gmtTime.length() == 14 && gmtDateTime.length() == 10) {
            // Remove the first 4 characters (YYYY) from GMT_time to get MMDDHHMMSS
            String gmtTimeStripped = gmtTime.substring(4);

            // Remove leading zeros from both strings for comparison
            String normalizedGmtDateTime = gmtDateTime.replaceFirst("^0+(?!$)", "");
            String normalizedGmtTimeStripped = gmtTimeStripped.replaceFirst("^0+(?!$)", "");

            // Compare the normalized values
            return normalizedGmtDateTime.equals(normalizedGmtTimeStripped);
        }
        return false;
    }


    // Validation 25 helper: Check if time_local (HHMMSS) matches TSTAMP_LOCAL (YYYYMMDDHHMMSS)
    public static boolean validateTimeLocal(String timeLocal, String tstampLocal) {
        if (tstampLocal.length() == 14 && timeLocal.length() == 6) {
            // Remove the first 8 characters (YYYYMMDD) from TSTAMP_LOCAL to get HHMMSS
            String tstampLocalStripped = tstampLocal.substring(8);

            // Remove leading zeros from both strings for comparison
            String normalizedTimeLocal = timeLocal.replaceFirst("^0+(?!$)", "");
            String normalizedTstampLocalStripped = tstampLocalStripped.replaceFirst("^0+(?!$)", "");

            // Compare the normalized values
            return normalizedTimeLocal.equals(normalizedTstampLocalStripped);
        }
        return false;
    }

    // Validation 26 helper: Check if date_local (MMDD) matches MMDD portion of TSTAMP_LOCAL (YYYYMMDDHHMMSS)
    public static boolean validateDateLocal(String dateLocal, String tstampLocal) {
        if (tstampLocal.length() >= 8 && dateLocal.length() == 4) {
            // Extract MMDD from TSTAMP_LOCAL
            String tstampLocalMMDD = tstampLocal.substring(4, 8);

            // Remove leading zeros from both strings for comparison
            String normalizedDateLocal = dateLocal.replaceFirst("^0+(?!$)", "");
            String normalizedTstampLocalMMDD = tstampLocalMMDD.replaceFirst("^0+(?!$)", "");

            // Compare the normalized values
            return normalizedDateLocal.equals(normalizedTstampLocalMMDD);
        }
        return false;
    }


    // Validation 28 helper: Check if merchant_type matches MERCH_TYPE
    public static boolean validateMerchantType(String merchantType, String merchType) {
        return merchantType.equals(merchType);
    }


    // Validation 30 helper: Check if card_acceptor_term_id, CARD_ACPT_TERM_ID, and NET_TERM_ID are the same
    public static boolean validateTerminalIds(String cardAcceptorTermId, String cardAcptTermId, String netTermId) {
        return cardAcceptorTermId.equals(cardAcptTermId) && cardAcptTermId.equals(netTermId);
    }


    // Validation 31 helper: Check if Card_acceptor_id_code matches CARD_ACPT_ID
    public static boolean validateCardAcceptorId(String cardAcceptorIdCode, String cardAcptId) {
        return cardAcceptorIdCode.equals(cardAcptId);
    }


    // Validation: Check if card_acceptor_name_loc matches CARD_ACPT_NAME_LOC
    public static boolean validateNameLoc(String cardAcceptorNameLoc, String cardAcptNameLoc) {
        return cardAcceptorNameLoc.equals(cardAcptNameLoc);
    }

    // Validation: Check if CARD_ACPT_COUNTRY is "AUS"
    public static boolean validateCountryCode(String cardAcptCountry, String expectedCountryCode) {
        return cardAcptCountry.equals(expectedCountryCode);
    }

    // Validation: Check if card_acceptor_id_code has the expected format
    public static boolean validateIdCodeFormat(String cardAcceptorIdCode, String expectedFormat) {
        return cardAcceptorIdCode.equals(expectedFormat);
    }


    // Validation 33 helper: Validate each field against pos_data_code
    public static boolean validatePosDataCodeMapping(String posDataCode, Row row, Map<String, Integer> columnMap) {
        if (posDataCode == null || posDataCode.length() < 14) {
            return false; // Ensure pos_data_code is long enough
        }

        // Map of positions to field names (adjust the order as per requirement)
        String[] fieldNames = {
                "POS_CRD_DAT_IN_CAP", "POS_CRDHLDR_AUTH_C", "POS_CARD_CAP_CAP", "POS_OPER_ENV",
                "POS_CRDHLDR_PRESNT", "POS_CARD_PRES", "POS_CRD_DAT_IN_MOD", "POS_CRDHLDR_A_METH",
                "POS_CRDHLDR_AUTH", "POS_CRD_DAT_OT_CAP", "POS_TERM_OUT_CAP", "POS_PIN_CAPT_CAP",
                "POS_TERM_OPTR", "TERM_CLASS"
        };

        // Iterate through the fields and validate against posDataCode positions
        for (int i = 0; i < fieldNames.length; i++) {
            String fieldName = fieldNames[i];
            int expectedValue = Character.getNumericValue(posDataCode.charAt(i)); // Get expected value from pos_data_code
            int actualValue = Integer.parseInt(getCellValueAsString(row.getCell(columnMap.get(fieldName))));

            if (actualValue != expectedValue) {
                return false; // Return false if any field does not match
            }
        }

        return true; // Return true if all fields match
    }


    // Validation 43 helper: Validate if ret_ref_no and RETRIEVAL_REF_NO are identical
    public static boolean validateRetRefNo(String retRefNo, String retrievalRefNo) {
        if (retRefNo == null || retrievalRefNo == null) {
            return false; // Return false if either value is null
        }
        // Compare the two values after trimming whitespace
        return retRefNo.trim().equals(retrievalRefNo.trim());
    }


    // Validation 45 helper: Validate Msg_reason_code_rev and MSG_RESON_CODE_ACQ for MTI = 1430
    public static boolean validateMsgReasonCodeForMti1430(String mti, String msgReasonCodeRev, String msgResonCodeAcq) {
        if ("1430".equals(mti)) {
            // Only perform validation if MTI is 1430
            return msgReasonCodeRev.trim().equals(msgResonCodeAcq.trim());
        }
        return true; // If MTI is not 1430, skip validation
    }


    // Placeholder: This map should contain actual decryption keys mapped by their key ID.
    private static final Map<String, SecretKeySpec> decryptionKeys = new HashMap<>();

    static {
        // Load your decryption keys based on key IDs
        // For example, decryptionKeys.put("03", new SecretKeySpec(yourKeyBytes, "DESede"));
    }

    // Function to decrypt the encrypted PAN using the key identifier
    public static String decryptPan(String encryptedPan) {
        if (encryptedPan == null || encryptedPan.length() < 2) {
            return null;
        }

        String keyId = encryptedPan.substring(0, 2);  // Extract the 2-digit key ID
        String encryptedData = encryptedPan.substring(2);  // Encrypted PAN data after key ID

        SecretKeySpec keySpec = decryptionKeys.get(keyId);
        if (keySpec == null) {
            System.out.println("No decryption key found for key ID: " + keyId);
            return null;
        }

        try {
            Cipher cipher = Cipher.getInstance("DESede/ECB/PKCS5Padding");
            cipher.init(Cipher.DECRYPT_MODE, keySpec);
            byte[] decryptedBytes = cipher.doFinal(Base64.getDecoder().decode(encryptedData));
            return new String(decryptedBytes);  // Convert decrypted bytes to String (assuming UTF-8)
        } catch (Exception e) {
            System.out.println("Error during decryption: " + e.getMessage());
            return null;
        }
    }

    // Method for Validation 42: Validate if decrypted PAN matches the PAN column
    public static boolean validateDecryptedPan(String panEncrypted, String pan) {
        String decryptedPan = decryptPan(panEncrypted);
        return decryptedPan != null && decryptedPan.equals(pan);
    }


    // Validation 41: Checks if Secure_3d_result is null and CAVV_RESULT is empty
    public static boolean validateSecure3DAndCAVVResult(String secure3dResult, String cavvResult) {
        return secure3dResult == null && (cavvResult == null || cavvResult.isEmpty());
    }

    // Validation 40: Checks if Card_verification_result is null and CVV_CVN_RESULT is empty
    public static boolean validateCardVerificationAndCVVResult(String cardVerificationResult, String cvvCvnResult) {
        return cardVerificationResult == null && (cvvCvnResult == null || cvvCvnResult.isEmpty());
    }


    // Validation 39: Checks if Snknode_acquiring_inst and INST_ID_ISS have the same value
    public static boolean validateSnknodeAcquiringInstAndInstIdIss(String snknodeAcquiringInst, String instIdIss) {
        return snknodeAcquiringInst != null && snknodeAcquiringInst.equals(instIdIss);
    }

}
