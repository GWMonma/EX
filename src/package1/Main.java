package a;

import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
 //実行するとsample.xlsxのシート0にHelloWorldと記入される
public class Main {
 
	   public static void main(String[]  args) {
		   
	        FileOutputStream fos = null;
	        XSSFWorkbook workbook = null;
	 
	        try {
	 
	            // ワークブック→シート→行→セルの生成
	            workbook = new XSSFWorkbook();
	            XSSFSheet sheet = workbook.createSheet();
	            //行　0からスタート
	            XSSFRow row = sheet.createRow(0);
	            //列　0からスタート
	            XSSFCell cell = row.createCell(0);
	 
	            // セルの書式の生成
	            XSSFCellStyle cellStyle = workbook.createCellStyle();
	            XSSFFont font = workbook.createFont();
	            font.setFontName("ＭＳ ゴシック");
	            cellStyle.setFont(font);
	            cell.setCellStyle(cellStyle);
	 
	            // セルに書き込み
	            cell.setCellValue("Hello World!");
	 
	            // ファイル書き込み
	            fos = new FileOutputStream("C:\\Users\\Education\\Documents\\研修\\販売管理システム開発\\sample1.xlsx");
	            workbook.write(fos);
	 
	        } catch(Exception e) {
	            e.printStackTrace();
	        } finally {
	             try {
	                 if (fos != null) {
	                     fos.close();
	                 }
	                 if (workbook != null) {
	                     workbook.close();
	                 }
	             } catch(Exception e) {
	                 e.printStackTrace();
	             }
	        }
	   }
}