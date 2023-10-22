package ReadExcel;

import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ReadExcelSheet {
	
	public static void main(String[] args) throws Exception {
		
		FileInputStream fileInputstream = new FileInputStream("C:\\Users\\hp\\eclipse-workspace\\ReadExcel\\ExcelFile\\Text.xls");
		Sheet sh = WorkbookFactory.create(fileInputstream).getSheet("Sheet0");
		
		String uname = sh.getRow(0).getCell(1).getStringCellValue();
		System.out.println(uname);
		
		for(int i=0; i<sh.getLastRowNum(); i++)
		{
			int lastCellNum = sh.getRow(i).getLastCellNum();
			for(int j=0; j<lastCellNum; j++)
			{
				CellType ctype = sh.getRow(i).getCell(j).getCellType();
				String values = "";
				double intValues = 0;
				if(ctype.toString().equalsIgnoreCase("string"))
				
					values = sh.getRow(i).getCell(j).getStringCellValue();
					
				
				else 
					intValues = sh.getRow(i).getCell(j).getNumericCellValue();
				System.out.println(values);
				System.out.println(intValues);
				
			}
		}
		
		
	}

}
