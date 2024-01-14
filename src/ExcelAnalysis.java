import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

public class ExcelAnalysis {

    public static void main(String[] args) {
        String filePath = "F:/Assignment_Timecard.xlsx";

        try (FileInputStream file = new FileInputStream(new File(filePath));
             Workbook workbook = new XSSFWorkbook(file)) {

            Sheet sheet = workbook.getSheetAt(0);

            // Process the sheet data
            analyzeSheet(sheet);

        } catch (IOException | ParseException e) {
            e.printStackTrace();
        }
    }

    private static void analyzeSheet(Sheet sheet) throws ParseException {
        // Your analysis logic goes here
        // Example: Print cell values

        // Assuming the data starts from row 1 (excluding header)
        Map<String, Boolean> processedIds = new HashMap<>();
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);

            String positionId = row.getCell(0).getStringCellValue();

            // Check if the position ID has already been processed
            if (!processedIds.containsKey(positionId)) {
                // Handle time cell values which may be either numeric or string
                Date timeIn = getCellDateValue(row.getCell(2));
                Date timeOut = getCellDateValue(row.getCell(3));

                // Additional analysis based on constraints
                if (hasWorkedFor7ConsecutiveDays(sheet, i)) {
                    System.out.println("Position ID: " + positionId + " has worked for 7 consecutive days.");
                }

                if (hasLessThan10HoursBetweenShifts(timeIn, timeOut)) {
                    System.out.println("Position ID: " + positionId + " has less than 10 hours between shifts.");
                }

                if (hasWorkedMoreThan14HoursInSingleShift(sheet, i)) {
                    System.out.println("Position ID: " + positionId + " has worked for more than 14 hours in a single shift.");
                }

                // Mark the position ID as processed
                processedIds.put(positionId, true);
            }
        }
    }

    private static boolean hasWorkedFor7ConsecutiveDays(Sheet sheet, int rowIndex) throws ParseException {
        if (rowIndex >= 7) {
            // Check if the employee has worked for 7 consecutive days
            // Example: compare the dates in the current row and previous 6 rows
            // Return true if the condition is met
            return true; // Placeholder, replace with your logic
        }
        return false;
    }

    private static boolean hasLessThan10HoursBetweenShifts(Date timeIn, Date timeOut) {
        if (timeIn != null && timeOut != null) {
            long timeDifferenceMillis = timeOut.getTime() - timeIn.getTime();
            long hoursDifference = timeDifferenceMillis / (60 * 60 * 1000);

            // Check if the hours difference is less than 10 but greater than 1
            return hoursDifference > 1 && hoursDifference < 10;
        }
        return false;
    }

    private static boolean hasWorkedMoreThan14HoursInSingleShift(Sheet sheet, int rowIndex) throws ParseException {
        // Your logic to check if an employee has worked for more than 14 hours in a single shift
        // Example: Compare the time in and time out values in the current row
        return true; // Placeholder, replace with your logic
    }

    // Method to handle both numeric and string date cell values
    private static Date getCellDateValue(Cell cell) throws ParseException {
        if (cell == null) {
            return null;
        }

        if (cell.getCellType() == CellType.NUMERIC) {
            // Handle numeric date value (assuming it's a date)
            return cell.getDateCellValue();
        } else if (cell.getCellType() == CellType.STRING) {
            // Handle string date value (assuming it's in a recognizable date format)
            return parseStringTime(cell.getStringCellValue());
        }

        return null;
    }

    // Method to parse time string with or without AM/PM
    private static Date parseStringTime(String timeStr) throws ParseException {
        if (timeStr == null || timeStr.trim().isEmpty()) {
            // Handle empty or invalid time string
            return null;
        }

        // Check if the time format includes AM/PM
        boolean hasAmPm = timeStr.toLowerCase().contains("am") || timeStr.toLowerCase().contains("pm");

        SimpleDateFormat timeFormat;

        if (hasAmPm) {
            // Handle time with AM/PM
            timeFormat = new SimpleDateFormat("hh:mm a");
        } else {
            // Handle time without AM/PM
            timeFormat = new SimpleDateFormat("HH:mm");
        }

        return timeFormat.parse(timeStr);
    }
}
