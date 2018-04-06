package svutest;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.util.Iterator;
import java.util.LinkedHashMap;

import javax.xml.transform.Result;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;

import oracle.sql.INTERVALYM;

public class MyThread {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		
		
		try {
			String url="jdbc:oracle:thin:@sun12.albertsons.com:1532:FIN1DEV1";
			String user="ABSN_READ";
			String password="dev1absnread";
			DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
			Connection conn=DriverManager.getConnection(url, user, password);
			
			
			PreparedStatement stmt=null;
			//Workbook
			HSSFWorkbook workBook=new HSSFWorkbook();
			HSSFSheet sheet1=null;
			
			//Cell
			Cell c=null;
			
			CellStyle cs=workBook.createCellStyle();
			HSSFFont f =workBook.createFont();
			f.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
			f.setFontHeightInPoints((short) 12);
	        cs.setFont(f);
	        
	        
	        sheet1=workBook.createSheet("Sheet1");
	        
	        
	        String query="";
	        stmt=conn.prepareStatement(query);
	        ResultSet rs=stmt.executeQuery();
	        
	        ResultSetMetaData metaData=rs.getMetaData();
	        int colCount=metaData.getColumnCount();
	        
	        LinkedHashMap<Integer, TableInfo> hashMap=new LinkedHashMap<Integer, TableInfo>();
	        
	        
	        for(int i=0;i<colCount;i++){
	        	TableInfo tableInfo=new TableInfo();
	        	tableInfo.setFieldName(metaData.getColumnName(i+1).trim());
	        	tableInfo.setFieldText(metaData.getColumnLabel(i+1));
	        	tableInfo.setFieldSize(metaData.getPrecision(i+1));
	        	tableInfo.setFieldDecimal(metaData.getScale(i+1));
	        	tableInfo.setFieldType(metaData.getColumnType(i+1));
	            tableInfo.setCellStyle(getCellAttributes(workBook, c, tableInfo));
	        	
	        	hashMap.put(i, tableInfo);
	        }
	    
			//Row and Column Indexes
	        int idx=0;
	        int idy=0;
	        
	        HSSFRow row=sheet1.createRow(idx);
	        TableInfo tableInfo=new TableInfo(); 
	        
	        Iterator<Integer> iterator=hashMap.keySet().iterator();
	        
	        while(iterator.hasNext()){
	        Integer key=(Integer)iterator.next();
	        
	        tableInfo=hashMap.get(key);
	        c=row.createCell(idy);
	        c.setCellValue(tableInfo.getFieldText());
	        c.setCellStyle(cs);
	        if(tableInfo.getFieldSize() > tableInfo.getFieldText().trim().length()){
                sheet1.setColumnWidth(idy, (tableInfo.getFieldSize()* 500));
            }
            else {
                sheet1.setColumnWidth(idy, (tableInfo.getFieldText().trim().length() * 500));
            }
            idy++;
	        }
	        
	        while (rs.next()) {
	               
                idx++;
                row = sheet1.createRow(idx);
                System.out.println(idx);
                for (int i = 0; i < colCount; i++) {

                    c = row.createCell(i);
                    tableInfo = hashMap.get(i);

                    switch (tableInfo.getFieldType()) {
                    case 1:
                        c.setCellValue(rs.getString(i+1));
                        break;
                    case 2:
                        c.setCellValue(rs.getDouble(i+1));
                        break;
                    case 3:
                        c.setCellValue(rs.getDouble(i+1));
                        break;
                    default:
                        c.setCellValue(rs.getString(i+1));
                        break;
                    }
                    c.setCellStyle(tableInfo.getCellStyle());
                }
       
            }
	        rs.close();
            stmt.close();
            conn.close();
            
            String path="";
            
            FileOutputStream fileOut = new FileOutputStream(path);

            workBook.write(fileOut);
            fileOut.close();

	        
	        
		
			
		} catch (SQLException | FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}
	
	private static CellStyle getCellAttributes (Workbook wb, Cell c, TableInfo db2TableInfo){
 
        CellStyle cs= wb.createCellStyle();
        DataFormat df = wb.createDataFormat();
        Font f = wb.createFont();

        switch (db2TableInfo.getFieldDecimal()) {
        case 1:
            cs.setDataFormat(df.getFormat("#,##0.0"));
            break;
        case 2:
            cs.setDataFormat(df.getFormat("#,##0.00"));
            break;
        case 3:
            cs.setDataFormat(df.getFormat("#,##0.000"));
            break;
        case 4:
            cs.setDataFormat(df.getFormat("#,##0.0000"));
            break;
        case 5:
            cs.setDataFormat(df.getFormat("#,##0.00000"));
            break;
        default:
            break;
        }

        cs.setFont(f);

        return cs;
	}

}
