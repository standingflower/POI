package useJxl;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Scanner;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class test {
	
	public static void main(String[] args) throws IOException {
		
		Scanner input=new Scanner(System.in);
		System.out.println("请输入姓名：");
		
		String name=input.nextLine();
		
		
		test.getFileName("D:\\桌面\\关于素拓积分卡的事\\关于素拓积分卡的事\\签到表 (2)\\签到表 (2)\\签到表",name);
	}

	//获取指定文件夹的所文件名称
	public static void getFileName(String filePath,String name) throws IOException{
		
		File f=new File(filePath);
		
		if(!f.exists()){
			System.out.println("路径不存在");
			
			return;
		}
		
		File fp[]=f.listFiles();
		
		for(int i=0;i<fp.length;i++){
			File fs=fp[i];
			
			if(fs.isDirectory()){
				System.out.println(fs.getName()+"目录");
			}else{
//				System.out.println(fs.getName());
				reader(fs.getName(),name);
			}
		}
		
		
	}
	
	private static void reader(String filePath,String name) throws IOException{
		InputStream stream = new FileInputStream("D:\\桌面\\关于素拓积分卡的事\\关于素拓积分卡的事\\签到表 (2)\\签到表 (2)\\签到表\\"+filePath);
		Workbook wb = null;
		wb = new XSSFWorkbook(stream);
		Sheet sheet1 = wb.getSheetAt(0);
		
		
		for (Row row : sheet1) {
			for (Cell cell : row) {
				if(cell!=null){
					cell.setCellType(Cell.CELL_TYPE_STRING);
					if(cell.getStringCellValue().equals(name)){
						System.out.println(filePath);
					}
				}
			}
//			System.out.println();
		}
	}
	
}
