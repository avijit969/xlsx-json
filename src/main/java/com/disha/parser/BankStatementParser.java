package com.disha.parser;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;
import com.fasterxml.jackson.datatype.jsr310.JavaTimeModule;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.*;

public class BankStatementParser {

    private static final Set<String> HEADER_KEYWORDS = new HashSet<>(Arrays.asList(
            "date", "transaction", "amount", "balance", "withdrawal",
            "deposit", "description", "details", "narration", "debit", "credit", "ref", "reference",
            "value date", "post date", "particulars", "cheque", "chk"
    ));

    public static void main(String[] args) {
        if (args.length < 2) {
            System.err.println("Usage: java com.disha.parser.BankStatementParser <bankName> <excelFilePath>");
            return;
        }

        String bankName = args[0];
        String filePath = args[1];

        try {
            List<Map<String, Object>> transactions = parseExcel(filePath);

            Map<String, Object> output = new LinkedHashMap<>();
            output.put("bankName", bankName);
            output.put("count", transactions.size());
            output.put("transactions", transactions);

            ObjectMapper mapper = new ObjectMapper();
            mapper.registerModule(new JavaTimeModule());
            mapper.configure(SerializationFeature.WRITE_DATES_AS_TIMESTAMPS, false);
            mapper.setDateFormat(new SimpleDateFormat("yyyy-MM-dd"));
            String jsonOutput = mapper.writerWithDefaultPrettyPrinter().writeValueAsString(output);

            System.out.println(jsonOutput);
            
            //save to a file
            String outFileName = filePath.substring(0, filePath.lastIndexOf(".")) + ".json";
            Files.write(Paths.get(outFileName), jsonOutput.getBytes());
            System.out.println("\nSuccessfully written output to " + outFileName);

        } catch (Exception e) {
            System.err.println("Error parsing file:");
            e.printStackTrace();
        }
    }

    public static List<Map<String, Object>> parseExcel(String filePath) throws IOException {
        List<Map<String, Object>> transactions = new ArrayList<>();

        try (Workbook workbook = WorkbookFactory.create(new File(filePath))) {
            // Take the first active sheet
            Sheet sheet = workbook.getSheetAt(workbook.getActiveSheetIndex() <= 0 ? 0 : workbook.getActiveSheetIndex());

            int maxScore = 0;
            int headerRowIndex = -1;
            Map<Integer, String> columns = new HashMap<>();

            // 1. Identify header row
            for (Row row : sheet) {
                int score = 0;
                Map<Integer, String> currentColumns = new HashMap<>();

                for (Cell cell : row) {
                    String cellValue = extractCellValueAsString(cell).toLowerCase().trim();
                    if (!cellValue.isEmpty()) {
                        currentColumns.put(cell.getColumnIndex(), cellValue);
                        // Check if it looks like a known transaction header
                        for (String kw : HEADER_KEYWORDS) {
                            if (cellValue.contains(kw) && cellValue.length() <= kw.length() + 15) {
                                score++;
                                break;
                            }
                        }
                    }
                }

                if (score > maxScore && score >= 2) {
                    maxScore = score;
                    headerRowIndex = row.getRowNum();
                    columns = currentColumns;
                }
            }

            if (headerRowIndex == -1) {
                System.err.println("Warning: Could not identify a clear header row. Trying row 0.");
                headerRowIndex = 0;
                Row r = sheet.getRow(0);
                if (r != null) {
                    for (Cell cell : r) {
                        columns.put(cell.getColumnIndex(), extractCellValueAsString(cell).trim());
                    }
                }
            }

            // 2. Parse transactions below the header
            DataFormatter dataFormatter = new DataFormatter();
            
            // Loop from the row immediately following the header
            for (int rIndex = headerRowIndex + 1; rIndex <= sheet.getLastRowNum(); rIndex++) {
                Row row = sheet.getRow(rIndex);
                if (row == null) {
                    continue; // Skip completely empty rows
                }

                Map<String, Object> transaction = new LinkedHashMap<>();
                boolean isEmptyRow = true;
                
                int consecutiveEmptyThreshold = columns.size() > 0 ? (columns.size() / 2) + 1 : 3;
                int consecutiveEmptyCells = 0;

                for (Map.Entry<Integer, String> col : columns.entrySet()) {
                    int colIndex = col.getKey();
                    String colName = col.getValue();
                    if (colName.isEmpty()) {
                        colName = "Column_" + colIndex;
                    }

                    Cell cell = row.getCell(colIndex, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                    if (cell != null) {
                        Object cellValue = extractCellValue(cell, dataFormatter);
                        if (cellValue != null && !cellValue.toString().trim().isEmpty()) {
                            transaction.put(colName, cellValue);
                            isEmptyRow = false;
                            consecutiveEmptyCells = 0;
                        } else {
                            consecutiveEmptyCells++;
                        }
                    } else {
                        consecutiveEmptyCells++;
                    }
                }

                // If the entire row was empty (or no mapped columns had data), we might have reached end of data
                if (isEmptyRow) {
                    // Usually there are intermediate empty rows, we could continue but ignore this one
                    continue; // You can also break if data format guarantees no empty rows in between
                }

                // Sometimes footers have one column filled with "Closing Balance". We could detect and stop, 
                // but just adding it as a transaction might be okay, or we could filter it out.
                // We'll keep it simple and just include if it has values.
                
                transactions.add(transaction);
            }
        }

        return transactions;
    }

    private static String extractCellValueAsString(Cell cell) {
        if (cell == null) return "";
        try {
            switch (cell.getCellType()) {
                case STRING:
                    return cell.getStringCellValue();
                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        return cell.getDateCellValue().toString();
                    }
                    return String.valueOf(cell.getNumericCellValue());
                case BOOLEAN:
                    return String.valueOf(cell.getBooleanCellValue());
                case FORMULA:
                    try {
                        return String.valueOf(cell.getNumericCellValue());
                    } catch (Exception e) {
                        return cell.getStringCellValue();
                    }
                case BLANK:
                default:
                    return "";
            }
        } catch(Exception e) {
            return "";
        }
    }

    private static Object extractCellValue(Cell cell, DataFormatter dataFormatter) {
        if (cell == null) return null;
        try {
            switch (cell.getCellType()) {
                case STRING:
                    return cell.getStringCellValue().trim();
                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        return cell.getDateCellValue(); // Jackson will format it
                    }
                    return cell.getNumericCellValue();
                case BOOLEAN:
                    return cell.getBooleanCellValue();
                case FORMULA:
                    switch (cell.getCachedFormulaResultType()) {
                        case NUMERIC:
                            if (DateUtil.isCellDateFormatted(cell)) {
                                return cell.getDateCellValue();
                            }
                            return cell.getNumericCellValue();
                        case STRING:
                            return cell.getStringCellValue().trim();
                        case BOOLEAN:
                            return cell.getBooleanCellValue();
                        default:
                            return null;
                    }
                case BLANK:
                default:
                    return null;
            }
        } catch(Exception e) {
            return dataFormatter.formatCellValue(cell);
        }
    }
}
