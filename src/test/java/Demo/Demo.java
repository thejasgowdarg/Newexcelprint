package Demo;

import java.io.FileInputStream;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Demo {

	public static void main(String[] args) throws Exception {
		
		String path="C:\\Users\\Tulasikumar\\eclipse-workspace\\Excel_Print\\src\\main\\java\\Book1.xlsx";
		String sheetname="Sheet1";
		
		FileInputStream file=new FileInputStream(path);
		XSSFWorkbook book=new XSSFWorkbook(file);
		XSSFSheet sheet=book.getSheet(sheetname);
		
		for(int i=0;i<sheet.getLastRowNum()+1;i++) {
			for(int j=0;j<sheet.getRow(0).getLastCellNum();j++) {
				String data=sheet.getRow(i).getCell(j).toString();
				System.out.print(" "+data);
			}
			System.out.println();
		}
		book.close();
	}
}

