package usePoi;

import java.io.File;
import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.util.CellRangeAddress;

import jxl.write.WritableWorkbook;

public class Write {

	public static final String url="jdbc:mysql://localhost:3306/test?useUnicode=true&characterEncoding=UTF-8";//连接数据库的URL地址
	public static final String username="root";//数据库的用户名
	public static final String password="zt1121576201";//数据库的密码
	
	public static Connection conn=null;//连接对象
	public static PreparedStatement stmt = null;
    public static ResultSet rs = null;//结果集
    
    File file;
    
    WritableWorkbook wwb;
    
    public static void main(String[] args) throws Exception {
		Write write = new Write();
		write.put(write.getMsg());
	}
	
    
    public Write(){
    	try {
			Class.forName("com.mysql.jdbc.Driver");
			if(conn==null){
				conn=DriverManager.getConnection(url,username,password);
			}
		} catch (ClassNotFoundException e) {
			System.out.println("驱动加载异常！！！");
			e.printStackTrace();
		} catch (SQLException e) {
			System.out.println("获取数据库链接异常");
			e.printStackTrace();
		}
    }
    
    
    public ResultSet getMsg() throws Exception{
		String sql="select * from user";
		
		stmt = conn.prepareStatement(sql);
		
		rs = stmt.executeQuery();
		
		System.out.println("列数"+rs.getMetaData().getColumnCount());
		
		return rs;
		
	}
    
    HSSFWorkbook workbook = null;
	
    POIFSFileSystem fs = null;
    
    public void put(ResultSet rs) throws Exception{
    	//输出流
    	FileOutputStream os= new FileOutputStream("D:\\桌面\\writqe.xls");
    	//创建一个Excel文件
    	workbook = new HSSFWorkbook();
    	/*
			HSSFWorkbook和XSSFWorkbook都是WorkBook的实现类
			但是HSSFWorkbook是对应的是xls
			XSSFWorkbook对应的是xlsx
    	 */
    	//页数
    	HSSFSheet sheet =null;
    	//行书
    	HSSFRow row = null;
    	//单元格
    	HSSFCell cell = null;
    	
    	int i=0;
    	//生成页
    	sheet = workbook.createSheet("第 一页");
    	//设置列的宽度
    	sheet.setColumnWidth(0, 4000);
    	
    	//设置字体
    	HSSFFont font = workbook.createFont();
    	//设置字体样式
    	font.setFontName("Verdana");
    	//设置加粗的磅数
    	font.setBoldweight((short) 100);
    	//设置字体的大小
    	font.setFontHeight((short) 300);
    	//设置字体的颜色
    	font.setColor(HSSFColor.BLUE.index);
    	
    	
    	
    	// 创建单元格样式
    	HSSFCellStyle style = workbook.createCellStyle();
    	
    	//设置左右居中
    	style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
    	//设置上下居中
    	style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
    	//设置前景色
    	style.setFillForegroundColor(HSSFColor.LIGHT_TURQUOISE.index);
    	//设置背景色
    	style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
    	// 设置单元格边框颜色
    	style.setBottomBorderColor(HSSFColor.RED.index);
    	//设置单元格的边框为粗体
    	style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
    	//设置左边框大小
    	style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
    	style.setBorderRight(HSSFCellStyle.BORDER_THIN);
    	style.setBorderTop(HSSFCellStyle.BORDER_THIN);
    	//将字体添加到单元格中
    	style.setFont(font);// 设置字体

    	// 合并单元格(startRow，endRow，startColumn，endColumn)
    	sheet.addMergedRegion(new CellRangeAddress(9, 15, 9, 15));
    	
    	//生成行
    	row = sheet.createRow(i);
    	//设置行高
    	row.setHeight((short) 500);
    	//生成单元格
    	cell = row.createCell(0,HSSFCell.CELL_TYPE_STRING);
    	cell.setCellValue("序号");
    	cell = row.createCell(1,HSSFCell.CELL_TYPE_STRING);
    	cell.setCellValue("姓名");
    	cell = row.createCell(2,HSSFCell.CELL_TYPE_STRING);
    	cell.setCellValue("性别");
    	cell = row.createCell(3,HSSFCell.CELL_TYPE_STRING);
    	cell.setCellValue("身高");
    	cell = row.createCell(4,HSSFCell.CELL_TYPE_STRING);
    	cell.setCellValue("体重");
    	cell = row.createCell(5,HSSFCell.CELL_TYPE_STRING);
    	cell.setCellValue("电话号码");
    	
    	// 给Excel的单元格设置样式和赋值
    	cell.setCellStyle(style);
    	
    	while(rs.next()){
    		row = sheet.createRow(i+1);
    		
    		//设置单元格内容的类型
    		cell = row.createCell(0,HSSFCell.CELL_TYPE_STRING);
    		cell.setCellValue(i+1);
    		cell = row.createCell(1,HSSFCell.CELL_TYPE_STRING);
    		cell.setCellValue(rs.getString(2));
    		cell = row.createCell(2,HSSFCell.CELL_TYPE_STRING);
    		cell.setCellValue(rs.getString(3));
    		cell = row.createCell(3,HSSFCell.CELL_TYPE_STRING);
    		cell.setCellValue(rs.getString(4));
    		cell = row.createCell(4,HSSFCell.CELL_TYPE_STRING);
    		cell.setCellValue(rs.getString(5));
    		cell = row.createCell(5,HSSFCell.CELL_TYPE_STRING);
    		cell.setCellValue(rs.getString(6));
    		i++;
    	}
    	
    	//将Excel作为输出流输出
    	workbook.write(os);
    	//关闭输出流
    	os.close();
    	
    }
    
	
}
