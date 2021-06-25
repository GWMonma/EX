package a;

import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//実行するとxlsxファイルが作成される
public class Main2 {
    static final String INPUT_DIR = "C:\\Users\\Education\\Documents\\研修\\販売管理システム開発\\";
 
    public static void main(String[] args) {
    	 Date dateObj = new Date();
    	 SimpleDateFormat format = new SimpleDateFormat( "yyyyMMdd_HHmmss" );
    	 String display = format.format( dateObj );
    	 
    	
        try {
        	
            Workbook wb = new XSSFWorkbook();
            FileOutputStream fileOut = new FileOutputStream(INPUT_DIR + "ClientOrder_" +display +".xlsx");
             
            String safeName = WorkbookUtil.createSafeSheetName("['aaa's test*?]");
            Sheet sheet1 = wb.createSheet(safeName);
             
            CreationHelper createHelper = wb.getCreationHelper();
             
            for(int i=0;i<5;i++) {
            //Rows(行)
            Row row = sheet1.createRow((short)i);
            //cell(列)
            Cell cell = row.createCell(0);
            cell.setCellValue(1);
 
            row.createCell(1).setCellValue(1.2);
            row.createCell(2).setCellValue(
                 createHelper.createRichTextString("sample string"));
            row.createCell(3).setCellValue(true);
            }
             
            wb.write(fileOut);
            fileOut.close();
 
             
        }catch (Exception e) {
            e.printStackTrace();
        } finally {
             
        }
        
       
 
    }
 

}
