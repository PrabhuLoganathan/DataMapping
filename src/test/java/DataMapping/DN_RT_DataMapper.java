package DataMapping;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import static DataMapping.DataMapperUtils.getColumnMap;

public class DN_RT_DataMapper {

    public static void main(String[] args) throws IOException {
        // Input file paths
        String sheet1Path = "/Users/prabhuloganathan/Downloads/FinalData/DN.xlsx";
        String sheet2Path = "/Users/prabhuloganathan/Downloads/FinalData/RT.xlsx";
        String outputPath = "/Users/prabhuloganathan/Downloads/FinalData/mappedDataSheet.xlsx";

        // Create a new workbook for the output file
        Workbook outputWorkbook = new XSSFWorkbook();

        // Read both Excel files
        try (FileInputStream sheet1Input = new FileInputStream(sheet1Path);
             FileInputStream sheet2Input = new FileInputStream(sheet2Path);
             Workbook workbook1 = new XSSFWorkbook(sheet1Input);
             Workbook workbook2 = new XSSFWorkbook(sheet2Input)) {

            System.out.println("Reading data from Excel sheets...");

            // Get the first sheets from both input files
            Sheet dnSheet = workbook1.getSheetAt(0);  // DN_Data.xlsx
            Sheet rtSheet = workbook2.getSheetAt(0);  // Realtime_Data.xlsx

            // Check if both sheets have data
            if (dnSheet.getPhysicalNumberOfRows() == 0 || rtSheet.getPhysicalNumberOfRows() == 0) {
                System.out.println("Error: One or both of the sheets are empty. Terminating program.");
                return;
            }

            // Step 1: Read headers and create column maps for both sheets
            Row rtHeaderRow = rtSheet.getRow(0);  // RT header row
            Row dnHeaderRow = dnSheet.getRow(0);  // DN header row

            Map<String, Integer> rtColumnMap = DataMapperUtils.getColumnMap(rtHeaderRow);
            Map<String, Integer> dnColumnMap = DataMapperUtils.getColumnMap(dnHeaderRow);

// Step 2: Create the combined header row in the output sheet
            Sheet outputMappingSheet = outputWorkbook.createSheet("MappedData");
            DataMapperUtils.createCombinedHeaderRow(outputMappingSheet, rtColumnMap, dnColumnMap);

// Step 3: Populate TransactionData map with data from DN and RT sheets
            Map<String, TransactionData> transactionDataMap = new HashMap<>();
            DataMapperUtils.populateTransactionDataMap(dnSheet, rtSheet, dnColumnMap, rtColumnMap, transactionDataMap);

// Step 4: Populate the output sheet with combined data from RT and DN sheets
           DataMapperUtils.populateCombinedOutputSheet(outputMappingSheet, transactionDataMap, rtColumnMap, dnColumnMap);


            DataMapperUtils.copySheetToOutput(outputWorkbook, dnSheet, "DN_Data");
            DataMapperUtils.copySheetToOutput(outputWorkbook, rtSheet, "Realtime_Data");

            // Save all sheets to the output file
            DataMapperUtils.writeOutputFile(outputWorkbook, outputPath);
        } catch (IOException e) {
            System.out.println("Error while reading the Excel files.");
            e.printStackTrace();
        }
    }

}
