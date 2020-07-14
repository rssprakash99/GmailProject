package one;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.format.CellFormatType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel_data {
	
	public ArrayList<String> getData(String testcaseName) throws IOException {
		ArrayList<String> a = new ArrayList<String>();
		String excelpath = System.getProperty("user.dir");
		System.out.println(excelpath);
		
		//Get access to Excel file
		FileInputStream fis = new FileInputStream(excelpath+"\\src\\test\\resources\\testdata\\testdata.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		
		//Get access to Excel sheet
		int sheets = workbook.getNumberOfSheets();
		for(int i=0; i<sheets; i++)
		{
			if(workbook.getSheetName(i).equalsIgnoreCase("datasheet"))
					{
				      XSSFSheet sheet = workbook.getSheetAt(i);
				      //Step-1: Identify Testcases column by scanning the entire 1st row
				      Iterator<Row> rows = sheet.iterator();
				      Row firstrow = rows.next();
				      
				      Iterator<Cell> ce = firstrow.cellIterator();
				      int k=0,column=0;
				      while(ce.hasNext())
				      {
				    	  Cell value = ce.next();
				    	  if(value.getStringCellValue().equalsIgnoreCase("TestData"))
				    	  {
				    		  //desired column
				    		  column=k;
				    	  }
				    	  k++;
				      }
				      System.out.println(column);
				      //once column is identified then scan entire testcase column to identify purchase testcase row.
				      while(rows.hasNext())
				      {
				    	  Row r = rows.next();
				    	  if(r.getCell(column).getStringCellValue().equalsIgnoreCase(testcaseName))
				    			  {
				    		        //after you grab purchase testcase row = pull all the data of that row and feed in to test.
				    		       Iterator<Cell> cv =r.cellIterator();
				    		       while(cv.hasNext())
				    		       {
				    		    	   Cell c = cv.next();
				    		    	   if(c.getCellTypeEnum()==CellType.STRING)
				    		    	   {
				    		    		   a.add(c.getStringCellValue());
				    		    	   }
				    		    	   else {
				    		    		   a.add(NumberToTextConverter.toText(c.getNumericCellValue()));
				    		    	   }
				    		    	   
				    		       }
				    			  }
				      }				      
				}
		}
		return a;

	}

	
}
