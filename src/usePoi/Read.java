package usePoi;

import java.io.File;
import java.io.FileInputStream;
import java.text.SimpleDateFormat;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Read {
	
	public static void main(String[] args) throws Exception{
		
		SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd");
		
		File file =new File("D:\\桌面\\write.xls");
		
		FileInputStream is = new FileInputStream(file);
		//两种创建方法都是一样的(看看源码就知道了)
		Workbook workbook = WorkbookFactory.create(file);
		//获取页数
		int sheets = workbook.getNumberOfSheets();
		
		for(int s=0;s<sheets;s++){
			//获取指定页
			Sheet sheet = workbook.getSheetAt(s);
			//指定页的行数
			int rows = sheet.getPhysicalNumberOfRows();
			//遍历行
			for(int r=0;r<rows;r++){
				//获取指定行
				Row row = sheet.getRow(r);
				//获取指定行的列数
				int cellNum = row.getPhysicalNumberOfCells();
				//遍历每一行
				for(int c=0;c<cellNum;c++){
					//获取行的一个单元格
					Cell cell = row.getCell(c);
					//获取类型
					int type = cell.getCellType();
					String cellValue=null;
					switch (type) {
					//文本
					case Cell.CELL_TYPE_STRING:
						cellValue=cell.getStringCellValue();	
						break;
					//数字，日期
					case Cell.CELL_TYPE_NUMERIC:
						if(DateUtil.isCellDateFormatted(cell)){
							cellValue=format.format(cell.getDateCellValue());//日期类型
						}else{
							cellValue=String.valueOf(cell.getNumericCellValue());//数字类型
						}
						break;
					//布尔类型
					case Cell.CELL_TYPE_BOOLEAN:
						cellValue=String.valueOf(cell.getBooleanCellValue());
						break;
					//空白
					case Cell.CELL_TYPE_BLANK:
						cellValue=cell.getStringCellValue();
						break;
					//错误
					case Cell.CELL_TYPE_ERROR:
						cellValue="错误1";
						break;
					//公式
					case Cell.CELL_TYPE_FORMULA:
						cellValue="错误2";
						break;
					default:
						cellValue="错误3";
					}
					System.out.print(cellValue+"   ");
				}
				System.out.println();
			}
			
		}
	}
	
	
}
