package practicejavapoi;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.*;

public class CreateWorkBook {

	public static void main(String[] args) throws IOException {
		
		/*
		 * XSSFWorkbook wb = new XSSFWorkbook();
		 * 
		 * FileOutputStream fos = new FileOutputStream("test.xlsx");
		 * 
		 * wb.write(fos); fos.close(); System.out.println("Wrote to file successfully");
		 */
		XSSFWorkbook wb;
		WorkBookUtility wbu = new WorkBookUtility();
		try {
			//wb = wbu.createWorkBook();	
			// wbu.openWorkbook();
			wbu.createSpreadsheet();
		}catch(IOException e) {
			System.out.println("File couldnot be created");
			e.printStackTrace();
		}
		
		
		

	}

}
