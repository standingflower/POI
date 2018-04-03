package usePoi;

import java.io.FileInputStream;
import java.text.SimpleDateFormat;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;

public class Read_01 {
	
	public static void main(String[] args) throws Exception {
		
		HSSFWorkbook wb=null;
		POIFSFileSystem fs=null;
		SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd");
		
		
		fs = new POIFSFileSystem(new FileInputStream("D:\\桌面\\关于素拓积分卡的事\\关于素拓积分卡的事\\签到表 (2)\\签到表 (2)\\签到表\\2016年11月26日四川省龙浩航校公费招飞行员.xlsx"));
		//创建Excel连接
		wb=new HSSFWorkbook(fs);
		//获取页
		HSSFSheet sheet = wb.getSheetAt(0);
		//获取页数
		int sheets = wb.getNumberOfSheets();
		
		String cellValue=null;
		//迭代每一页的内容
		//每页中的行迭代器
		Iterator<Row> iterator = sheet.rowIterator();
		while(iterator.hasNext()){
			Row row = iterator.next();
			//每行中的单元格
			Iterator<Cell> iterator2 = row.iterator();
			
			while(iterator2.hasNext()){
				Cell cell = iterator2.next();
				
				//获取单元格内容类型
				int type = cell.getCellType();
				
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
