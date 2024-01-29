package excelprogram;

import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelCode1 {
	static FileInputStream f;
	static XSSFWorkbook w;
	static XSSFSheet sh;
	public static String getStringData(int a,int b) throws IOException //to get string 
	{
		f = new FileInputStream("C:\\Users\\Vismitha\\OneDrive\\Desktop\\JavaExample book.xlsx");//give file location-file open
		w=new XSSFWorkbook(f);
		sh=w.getSheet("Sheet1");
		Row r=sh.getRow(a);//row interface
		Cell c=r.getCell(b);//cell interface
		return c.getStringCellValue();//to get string
		
	}
	public static String getIntegerData(int a,int b) throws IOException 
	{
		f = new FileInputStream("C:\\\\Users\\\\Vismitha\\\\OneDrive\\\\Desktop\\\\JavaExample book.xlsx");
		w=new XSSFWorkbook(f);
		sh=w.getSheet("Sheet1");
		Row r=sh.getRow(a);
		Cell c=r.getCell(b);
		int x=(int) c.getNumericCellValue();
		return String.valueOf(x);//to return string having integer value
		
		
	}
}

