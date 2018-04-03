package useJxl;


import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;  
import java.io.IOException;  
import java.io.InputStream;  
import java.io.OutputStream;  
  
import org.apache.poi.hssf.usermodel.HSSFWorkbook;  
import org.apache.poi.ss.usermodel.Cell;  
import org.apache.poi.ss.usermodel.Row;  
import org.apache.poi.ss.usermodel.Sheet;  
import org.apache.poi.ss.usermodel.Workbook;  
import org.apache.poi.xssf.usermodel.XSSFWorkbook;  

//File file=new File("D:\\桌面\\关于素拓积分卡的事\\关于素拓积分卡的事\\签到表 (2)\\签到表 (2)\\签到表\\2016年11月26日四川省龙浩航校公费招飞行员.xlsx");
public class test01 {
	public static void main(String[] args) throws IOException {
		
		reader("D:\\桌面\\关于素拓积分卡的事\\关于素拓积分卡的事\\签到表 (2)\\签到表 (2)\\签到表\\职前培训.xlsx");
	}

	private static void reader(String filePath) throws IOException{
		InputStream stream = new FileInputStream(filePath);
		Workbook wb = null;
		wb = new XSSFWorkbook(stream);
		Sheet sheet1 = wb.getSheetAt(0);
		
		
		for (Row row : sheet1) {
			for (Cell cell : row) {
				if(cell!=null){
					cell.setCellType(Cell.CELL_TYPE_STRING);
					System.out.print(cell.getStringCellValue() + "  ");
				}
				
				
			}
			System.out.println();
		}
	}
	
	
}
