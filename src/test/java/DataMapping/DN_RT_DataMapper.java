package DataMapping;


import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Map;

public class DN_RT_DataMapper {

    public static void main(String[] args) throws IOException {
        // Input file paths
        String sheet1Path = "/Users/prabhuloganathan/IdeaProjects/DataMapping/src/test/testData/DATA NAVIGATOR-3.xlsx";
        String sheet2Path = "/Users/prabhuloganathan/IdeaProjects/DataMapping/src/test/testData/Realtime-3.xlsx";
        String outputPath = "/Users/prabhuloganathan/IdeaProjects/DataMapping/src/test/testData/mappedDataSheet.xlsx";

        // Create a new workbook for the output file
        Workbook outputWorkbook = new XSSFWorkbook();

        // Read both Excel files
        try (FileInputStream sheet1Input = new FileInputStream(sheet1Path);
             FileInputStream sheet2Input = new FileInputStream(sheet2Path);
             Workbook workbook1 = new XSSFWorkbook(sheet1Input);
             Workbook workbook2 = new XSSFWorkbook(sheet2Input)) {


            System.out.println("Reading data from Excel sheets...");

            // Get the first sheets from both input files
            Sheet sheet1 = workbook1.getSheetAt(0);  // DN_Data.xlsx
            Sheet sheet2 = workbook2.getSheetAt(0);  // Realtime_Data.xlsx

            // Check if both sheets have data
            if (sheet1.getPhysicalNumberOfRows() == 0 || sheet2.getPhysicalNumberOfRows() == 0) {
                System.out.println("Error: One or both of the sheets are empty. Terminating program.");
                return;
            }

            // Get header rows from both sheets
            Row sheet1HeaderRow = sheet1.getRow(0);  // Header row for DN_Data
            Row sheet2HeaderRow = sheet2.getRow(0);  // Header row for Realtime_Data

            // Create column maps for both sheets
            Map<String, Integer> sheet1ColumnMap = DataMapperUtils.getColumnMap(sheet1HeaderRow);  // DN_Data columns
            Map<String, Integer> sheet2ColumnMap = DataMapperUtils.getColumnMap(sheet2HeaderRow);  // Realtime_Data columns

            // Map to store transaction numbers from DN_Data (Sheet1)
            Map<String, Row> dnDataMap = DataMapperUtils.createDataMap(sheet1, sheet1ColumnMap);  // Using TRAN_NUMBER from DN_Data


            // Create a header row for the output sheet
            Sheet outputMappingSheet = outputWorkbook.createSheet("MappedData");
            DataMapperUtils.createHeaderRow(outputMappingSheet);

            // Process the realtime data sheet and match transactions with DN data using column names
            DataMapperUtils.processRealtimeSheet(sheet2, dnDataMap, outputMappingSheet, sheet2ColumnMap, sheet1ColumnMap);

            // Copy original DN_Data and Realtime_Data sheets to the output file
            DataMapperUtils.copySheetToOutput(outputWorkbook, sheet1, "DN_Data");
            DataMapperUtils.copySheetToOutput(outputWorkbook, sheet2, "Realtime_Data");

            // Save all sheets to the output file
            DataMapperUtils.writeOutputFile(outputWorkbook, outputPath);
        } catch (IOException e) {
            System.out.println("Error while reading the Excel files.");
            e.printStackTrace();
        }
    }
}
