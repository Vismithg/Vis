package excelprogram;
import java.io.IOException;

public class ExcelMain {
	public static void main(String[] args) throws IOException
	{
	String d=ExcelCode1.getStringData(3, 2);
	System.out.println(d);
	String e=ExcelCode1.getIntegerData(3, 1);
	System.out.println(e);

	}
}