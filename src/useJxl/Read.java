package useJxl;

import java.io.File;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;

public class Read {
	
	public static void main(String[] args) throws Exception {
		File file=new File("D:\\桌面\\test.xls");
//		File file=new File("D:\\桌面\\关于素拓积分卡的事\\关于素拓积分卡的事\\签到表 (2)\\签到表 (2)\\签到表\\2016年11月26日四川省龙浩航校公费招飞行员.xlsx");
		//与Excel文件建立连接
		Workbook wb=Workbook.getWorkbook(file);
		//获取Excel的页数
		Sheet[] sheets =wb.getSheets();
		//遍历页数
		for (Sheet s : sheets) {
			System.out.println(s.getName()+" : ");
			//获取每页的行数
			int rows=s.getRows();
			//遍历每一行
			for(int i=0;i<rows;i++){
				System.out.print("行"+i+" : ");
				//获取每一行中的元素
				Cell[] cells=s.getRow(i);
				if(cells.length>0){
					//遍历每行的元素
					for (Cell c : cells) {
						//trim去掉字符串两端的空格
						String str=c.getContents().trim();
						System.out.print(str+";");
					}
					System.out.println();
				}
			}
		}
	}
	
}
