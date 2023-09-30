package com;

import java.io.FileInputStream;

import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelOperations {

    public static void main(String[] args) throws IOException {
        String filePath = "C:\\Users\\bansh\\OneDrive\\Desktop\\Project\\Assignment\\DataFile\\Assignment_Timecard.xlsx";
        FileInputStream inputStream = new FileInputStream(filePath);
        
        
        XSSFWorkbook wb = new XSSFWorkbook(inputStream);
        XSSFSheet sheet = wb.getSheetAt(0);

        int rowCount = sheet.getLastRowNum();
        int colCount = sheet.getRow(0).getLastCellNum();

        // Assuming the first row contains headers (name, position, date, etc.)
        int nameColumnIndex = -1;
        int dateColumnIndex = -1;

        // Find the column indices for "Employee Name" and "Date" headers
        XSSFRow headerRow = sheet.getRow(0);
        for (int j = 0; j < colCount; j++) {
            XSSFCell headerCell = headerRow.getCell(j);
            String headerCellValue = headerCell.getStringCellValue().trim();
       if (headerCellValue.equalsIgnoreCase("Employee Name")) {
           nameColumnIndex = j;
       } else if (headerCellValue.equalsIgnoreCase("Pay Cycle Start Date")) {
           dateColumnIndex = j;
       }
   }
        
        
       // Iterates over the records in the file and prints the name and position of employees 
       // a) who has worked for 7 consecutive days.  package com;
        
        
        
      //Check if the columns were found
/*   if (nameColumnIndex >= 0 && dateColumnIndex >= 0) {
         String currentEmployee = null;
         int consecutiveDays = 0;

         // Track consecutive working days for each employee
         Map<String, Integer> consecutiveDaysMap = new HashMap<String, Integer>();

         // Iterate over the records
         for (int i = 1; i <= rowCount; i++) { // Start from 1 to skip the header row
             XSSFRow row = sheet.getRow(i);
             XSSFCell nameCell = row.getCell(nameColumnIndex);
             XSSFCell dateCell = row.getCell(dateColumnIndex);

             String name = nameCell.getStringCellValue();
                Date date = null; // Initialize date as null

                // Check if dateCell is not null and is of type NUMERIC
                if (dateCell != null && dateCell.getCellType() == CellType.NUMERIC) {
                    date = dateCell.getDateCellValue();
                }

                // Check if it's the same employee as before and date is not null
                if (!name.equals(currentEmployee) && date != null) {
                    currentEmployee = name;
                    consecutiveDays = 1;
                } else if (date != null) {
                    // Check if the current date is consecutive to the previous date
                    if (isConsecutiveDay(date, dateCell, sheet)) {
                        consecutiveDays++;
                    } else {
                        consecutiveDays = 1;
                    }
                }

                // Update consecutive days in the map for the employee
                consecutiveDaysMap.put(name, consecutiveDays);
            }

                 for (Map.Entry<String, Integer> entry : consecutiveDaysMap.entrySet()) {
                System.out.println("Employee Name: " + entry.getKey() + ", Consecutive Days: " + entry.getValue());
            }
        } 

        wb.close();
        inputStream.close();
    }

    // Helper method to check if a date is consecutive to the previous date
    private static boolean isConsecutiveDay(Date currentDate, XSSFCell dateCell, XSSFSheet sheet) {
        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(currentDate);
        calendar.add(Calendar.DAY_OF_YEAR, -1); // Subtract 1 day
        Date previousDate = calendar.getTime();
       // Print employees who have worked for 7 consecutive days
//            System.out.println("Contents of consecutiveDaysMap:");

        // Check if the previous date exists in the sheet
        for (int i = dateCell.getRowIndex() - 1; i > 0; i--) {
            XSSFRow previousRow = sheet.getRow(i);
            XSSFCell previousDateCell = previousRow.getCell(dateCell.getColumnIndex());

            if (previousDateCell != null && previousDateCell.getCellType() == CellType.NUMERIC) {
                Date sheetDate = previousDateCell.getDateCellValue();
                if (dateFormat.format(sheetDate).equals(dateFormat.format(previousDate))) {
                    return true;
                }
            }
        }

        return false;
    }

      
    }

*/


//b)who have less than 10 hours of time between shifts but greater than 1 hour write a code for this?
        

        // Check if the columns were found
        if (nameColumnIndex >= 0 && dateColumnIndex >= 0) {
            String currentEmployee = null;
            Date previousDate = null;

            // Iterate over the records
            for (int i = 1; i <= rowCount; i++) { // Start from 1 to skip the header row
                XSSFRow row = sheet.getRow(i);
                XSSFCell nameCell = row.getCell(nameColumnIndex);
                XSSFCell dateCell = row.getCell(dateColumnIndex);

                String name = nameCell.getStringCellValue();
                Date currentDate = null; // Initialize date as null

                // Check if dateCell is not null and is of type NUMERIC
                if (dateCell != null && dateCell.getCellType() == CellType.NUMERIC) {
                    currentDate = dateCell.getDateCellValue();
                }

                // Check if it's the same employee as before and dates are not null
                if (name.equals(currentEmployee) && currentDate != null && previousDate != null) {
                    long timeDifference = currentDate.getTime() - previousDate.getTime();
                    long hoursDifference = timeDifference / (60 * 60 * 1000); // Convert to hours

                    // Check if the time difference is less than 10 hours but greater than 1 hour
                    if (hoursDifference > 1 && hoursDifference < 10) {
                        System.out.println("Employee Name: " + name);
                    }
                }

                // Update currentEmployee and previousDate
                currentEmployee = name;
                previousDate = currentDate;
                System.out.println("Name: " + name);
                System.out.println("Current Date: " + currentDate);
                System.out.println("Previous Date: " + previousDate);

            }
        }

        wb.close();
        inputStream.close();
    }
}


//c)Who has worked for more than 14 hours in a single shift         

        // Check if the columns were found
     /*   if (nameColumnIndex >= 0 && shiftStartColumnIndex >= 0 && shiftEndColumnIndex >= 0) {
            // Iterate over the records
            for (int i = 1; i <= rowCount; i++) { // Start from 1 to skip the header row
                XSSFRow row = sheet.getRow(i);
                XSSFCell nameCell = row.getCell(nameColumnIndex);
                XSSFCell shiftStartCell = row.getCell(shiftStartColumnIndex);
                XSSFCell shiftEndCell = row.getCell(shiftEndColumnIndex);

                String name = nameCell.getStringCellValue();
                String shiftStartTime = shiftStartCell.getStringCellValue();
                String shiftEndTime = shiftEndCell.getStringCellValue();

                // Parse shift start and end times
                String[] startTimeParts = shiftStartTime.split(":");
                String[] endTimeParts = shiftEndTime.split(":");

                int startHour = Integer.parseInt(startTimeParts[0]);
                int startMinute = Integer.parseInt(startTimeParts[1]);

                int endHour = Integer.parseInt(endTimeParts[0]);
                int endMinute = Integer.parseInt(endTimeParts[1]);

                // Calculate total minutes worked in the shift
                int totalMinutesWorked = (endHour - startHour) * 60 + (endMinute - startMinute);

                // Check if the total minutes worked is more than 840 (14 hours)
                if (totalMinutesWorked > 840) {
                    System.out.println("Employee Name: " + name);
                }
            }
        }

        wb.close();
        inputStream.close();
    }
}*/



       

     