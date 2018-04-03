package useJxl;

import java.io.File;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;

import jxl.Workbook;
import jxl.format.VerticalAlignment;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WritableCellFormat;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

public class ReadFromMySQL {
	public static final String url="jdbc:mysql://localhost:3306/test?useUnicode=true&characterEncoding=UTF-8";//连接数据库的URL地址
	public static final String username="root";//数据库的用户名
	public static final String password="zt1121576201";//数据库的密码
	
	public static Connection conn=null;//连接对象
	public static PreparedStatement stmt = null;
    public static ResultSet rs = null;//结果集
    
    File file;
    
    WritableWorkbook wwb;
	
	
	public static void mdain(String[] args) throws Exception {
		ReadFromMySQL test = new ReadFromMySQL();
		
		test.write(test.getMsg());
	}
	
	public ReadFromMySQL(){
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
	
	
	public void write(ResultSet rs) throws Exception{
		
		file=new File("D:\\桌面\\write.xls");
		//与Excel文件建立连接
		wwb = Workbook.createWorkbook(file);
		//设置指定也的名称
		WritableSheet ws = wwb.createSheet("第一页", 0);
		int i=1;
		//获取数据库的列数
		int row=rs.getMetaData().getColumnCount();
		
//		WritableCellFormat format = new WritableCellFormat();
//		
////		format.setAlignment(Alignment.CENTRE);
//		
//		format.setVerticalAlignment(VerticalAlignment.CENTRE); 
		
		//设置单元格的合并
		//索引列，索引行，索引列+合并的列数，索引行+合并的行数
		ws.mergeCells(10, 10, 10+5, 10+9);
		//设置列的宽度
		ws.setColumnView(0, "序号".length()+10);
		
		
		ws.addCell(new Label(0,0,"序号"));
		ws.addCell(new Label(1,0,"姓名"));
		ws.addCell(new Label(2,0,"性别"));
		ws.addCell(new Label(3,0,"身高"));
		ws.addCell(new Label(4,0,"体重"));
		ws.addCell(new Label(5,0,"电话号码"));
		
		
		Label label;
		
		while(rs.next()){
			
			for(int j=1;j<=row;j++){
				
				label=new Label(j-1,i,rs.getString(j));
				
				WritableCellFormat format = new WritableCellFormat();
				
				format.setVerticalAlignment(VerticalAlignment.CENTRE); 
					
				label.setCellFormat(format);
				
				ws.addCell(label);
			}
			i++;
		}
		
		//Number用于添加数字
		ws.addCell(new Number(9,9,1321.31));
		
		wwb.write();
		wwb.close();
	}
}
