package Excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Properties;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel {

	FileInputStream read;
	XSSFWorkbook wb;
	XSSFSheet sheet1;
	
	public Excel() throws IOException {
		
		FileInputStream ip;
		Properties prop=new Properties();
		ip=new FileInputStream("C:\\Users\\701288\\eclipse-workspace\\Timesheet\\Properties\\Config.properties");
		prop.load(ip);
		
		//Reading data from excel
		String File=prop.getProperty("excel");
        read=new FileInputStream(File);
        wb=new XSSFWorkbook(read);
        wb.close();
		
		
	}
	public String getdata(int sheet2,int row,int column)  
	{
		
	    
	
   sheet1=wb.getSheetAt(sheet2);
   String result=sheet1.getRow(row).getCell(column).getStringCellValue();
return result;
	}

}
