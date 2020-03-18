package practicejavapoi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WorkBookUtility {
	
	public XSSFWorkbook createWorkBook() throws IOException {
		
		//try {
			
		XSSFWorkbook wb = new XSSFWorkbook();
		
		FileOutputStream fos = new FileOutputStream("test1.xlsx");
		
		wb.write(fos);
		fos.close();
		System.out.println("Wrote to file successfully");
		return wb;
		/*}catch(IOException e) {
			System.out.println("Couldnot create file..");
			System.out.println(e.getStackTrace());
		}*/
	}
	
	public void openWorkbook() throws IOException {
		File f = new File("test.xlsx");
		FileInputStream fis = new FileInputStream(f);
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		
		if(f.exists() && f.isFile()) {
			System.out.println("File opened successfully");
		}
		else {
			System.out.println("File couldnot be opened");
		}
		
	}
	
	public void createSpreadsheet() throws IOException{
		FileOutputStream fos = new FileOutputStream("test.xlsx");
		XSSFWorkbook wb = new XSSFWorkbook();
		wb.createSheet("First sheet created");
		//wb.write(fos);
		XSSFSheet ws = wb.getSheet("First sheet created");
		XSSFRow wr = ws.createRow(1);
		wb.createSheet("First row created");
		XSSFCell wc = wr.createCell(1, 1);
		wb.createSheet("First Cell created");
		wc.setCellValue("This is first cell");
		wb.write(fos);
		fos.close();
		System.out.println("Created Sheet Successfully");
	}
	

}
