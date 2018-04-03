package useJxl;

import java.io.File;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

public class Write {
	
	public static void main(String[] args) throws Exception {
		File file;
		
		WritableWorkbook wwb;
		
		file=new File("D:\\桌面\\write.xls");
		//与Excel文件建立连接
		wwb = Workbook.createWorkbook(file);
		//设置指定也的名称
		WritableSheet ws = wwb.createSheet("页1", 0);
		
		//将其放入到指定位置
		ws.addCell(new Label(0,0,"0行0列"));
		ws.addCell(new Label(1,0,"0行1列"));
		ws.addCell(new Label(0,1,"1行0列"));
		ws.addCell(new Label(1,1,"1行1列"));
		//写入大Excel中
		wwb.write();
		//关闭连接
		wwb.close();
	}
	
}
