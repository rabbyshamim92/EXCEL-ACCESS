import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class readExcel {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		
		File xlFile = new File("C:\\Users\\rabby\\OneDrive\\Desktop\\Prosmart class\\book1 (1).xlsx");//shift + right click
		FileInputStream xlFIS = new FileInputStream(xlFile);// jarfile =https://poi.apache.org/download.html(zipfile [2nd file] is for windows)
		
		//Xls = HSSF, xlsx = XSSF 
		XSSFWorkbook xlbook = new XSSFWorkbook(xlFIS);
		XSSFSheet xlsheet =xlbook.getSheet("sheet1");
		
		XSSFRow xlrow =xlsheet.getRow(0);
		XSSFCell xlcell = xlrow.getCell(0);
		
		System.out.println(xlcell.getStringCellValue());//only access value text
		
		xlcell = xlrow.getCell(1);
		
		System.out.println(xlcell.getStringCellValue());//only access value text
		
		xlrow =  xlsheet.getRow(2);
		xlcell = xlrow.getCell(1);
		
		
		System.out.println(xlcell.getNumericCellValue());//only access numbers
		
		
		
//jarfile =https://poi.apache.org/download.html(zipfile( under Binary distribution) [2nd file] is for windows)
// copy lib,ooxmllib and last 8jar after notice folder 
	
	}


}
