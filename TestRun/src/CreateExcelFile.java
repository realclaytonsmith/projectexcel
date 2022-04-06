import java.io.File;
import java.io.FileOutputStream;
import java.util.Map;
import java.util.Scanner;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CreateExcelFile {
	 static Scanner in = new Scanner(System.in);
    public static void main(String[] args)
    {
    	System.out.println(" Press Enter To Start");
    	in.nextLine();
        // this is the blank workbook.
        XSSFWorkbook workbook = new XSSFWorkbook();
 
        // this creates a blank excel sheet.
        XSSFSheet sheet
            = workbook.createSheet("student Details");
 
        // I am creating an empty tree map of string and object.
        Map<String, Object[]> data
            = new TreeMap<String, Object[]>();
 
        // typing data to the object and using the put method.
        data.put("001",
                 new Object[] { "ID", "NAME", "LASTNAME" });
        data.put("002",
                 new Object[] { 001, "Brock", "Johnson" });
        data.put("003",
                 new Object[] { 002, "Trenton", "Wilson" });
        data.put("004", new Object[] { 003, "Drake", "Henderson" });
        data.put("005", new Object[] { 004, "Malik", "Craton" });
 
        // moving the data and typing it to a sheet
        Set<String> keyset = data.keySet();
 
        int rownum = 0;
 
        for (String key : keyset) {
 
            // this creates a new row to the sheet
            Row row = sheet.createRow(rownum++);
 
            Object[] objArr = data.get(key);
 
            int cellnum = 0;
 
            for (Object obj : objArr) {
 
                // this line creates a cell at the next column in that row
                Cell cell = row.createCell(cellnum++);
 
                if (obj instanceof String)
                    cell.setCellValue((String)obj);
 
                else if (obj instanceof Integer)
                    cell.setCellValue((Integer)obj);
            }
        }
 
        try {
 
            // writing my workbook
            FileOutputStream out = new FileOutputStream(
                new File("myclassroster.xlsx"));
            workbook.write(out);
 
            // this closes the output connection
            out.close();
 
            //message for successful run
            System.out.println(" The Student Roster Was Created Successfully! ");
        }
        // handles exceptions
        catch (Exception e) {
 
            // This is the print trace method
            e.printStackTrace();
        }
    }
}
