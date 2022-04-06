import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadFile {
	 public static void main(String[] args)
	    {
	 
	        // Using try to check exceptions
	        try {
	 
	            // this reads the file from the directory
	            FileInputStream file = new FileInputStream(
	                new File("myclassroster.xlsx"));
	 
	            // creating an instance referencing to the xlsx file
	            XSSFWorkbook workbook = new XSSFWorkbook(file);
	 
	            // Getting the first sheet from the workbook
	            XSSFSheet sheet = workbook.getSheetAt(0);
	 
	            // going thru one by one
	            Iterator<Row> rowIterator = sheet.iterator();
	 
	            while (rowIterator.hasNext()) {
	 
	                Row row = rowIterator.next();
	 
	                // in each row go one by one thru the columns
	                Iterator<Cell> cellIterator
	                    = row.cellIterator();
	 
	                while (cellIterator.hasNext()) {
	 
	                    Cell cell = cellIterator.next();
	 
	       
	                    switch (cell.getCellType()) {
	 
	                    
	                    case Cell.CELL_TYPE_NUMERIC:
	                        System.out.print(
	                            cell.getNumericCellValue()
	                            + "t");
	                        break;
	 
	                    
	                    case Cell.CELL_TYPE_STRING:
	                        System.out.print(
	                            cell.getStringCellValue()
	                            + "t");
	                        break;
	                    }
	                }
	 
	                System.out.println("");
	            }
	 
	            //closing the file
	            file.close();
	        }
	 
	        // handling the exceptions
	        catch (Exception e) {
	 
	            //print trace method
	            e.printStackTrace();
	        }
	    }
}
