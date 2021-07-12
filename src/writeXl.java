import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import javax.imageio.stream.FileImageInputStream;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class writeXl {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		File xlFile = new File("C:\\Users\\rabby\\OneDrive\\Desktop\\Prosmart class\\book1 (1).xlsx");//shift + right click
		FileInputStream xlFIS = new FileInputStream(xlFile);// jarfile =https://poi.apache.org/download.html(zipfile [2nd file] is for windows)
		
		XSSFWorkbook xlbook = new XSSFWorkbook(xlFIS);
		XSSFSheet xlsheet =xlbook.getSheet("sheet1");
		
		XSSFRow xlrow =xlsheet.getRow(0);
		FileOutputStream xlFos = new FileOutputStream(xlFile);
		xlrow.createCell(3).setCellValue("pass");
		xlbook.write(xlFos);
	}

}
