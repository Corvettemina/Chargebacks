package poi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Format extends convert {

	public static void Fill (File file, int num) throws IOException {
	FileInputStream in = new FileInputStream(file);
    Workbook wb = new XSSFWorkbook(in);
    Sheet sheet1 = wb.getSheetAt(num);
    
    XSSFFont font= (XSSFFont) wb.createFont();
    font.setFontHeightInPoints((short)11);
    font.setFontName("Calibri");
    font.setBold(true);
    
    for(int i = 0; i <2; i ++) {
    Row row1 = sheet1.getRow(i);
    CellStyle style = wb.createCellStyle();
    if(i ==0) {
    style.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());}
    
    if(i == 1) { 
    style.setFillForegroundColor(IndexedColors.SKY_BLUE.getIndex());	
    } 
    	
    style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    	for(int k = 0; k<17; k++) {
    	Cell cell = row1.getCell(k);
    	style.setFont(font);
    	cell.setCellStyle(style);
    	}
	}
    
	FileOutputStream out = new FileOutputStream(file);
    
	wb.write(out);
	wb.close();
	out.close();
	in.close();
	
	}
	
	public static void SizeCols (File input ) throws IOException {
	
		FileInputStream in = new FileInputStream(input);
		Workbook wb = new XSSFWorkbook(in);	
		
		CellStyle style = wb.createCellStyle();
		style.setBorderBottom(BorderStyle.MEDIUM);
		style.setBorderRight(BorderStyle.MEDIUM);
		style.setBorderLeft(BorderStyle.MEDIUM);
		
		XSSFFont font= (XSSFFont) wb.createFont();
	    font.setFontHeightInPoints((short)8);
	    font.setFontName("Arial");
	    style.setFont(font);
	    
		for(int i = 0; i < wb.getNumberOfSheets(); i ++) {
		Sheet sheet1 = wb.getSheetAt(i);
		Row r = sheet1.getRow(0);
		
		Iterator<Cell> ci = r.cellIterator();
		
		while(ci.hasNext()) {
			Cell cell = ci.next();
			sheet1.autoSizeColumn(cell.getColumnIndex());
			cell.setCellStyle(style);
			}
		}

		FileOutputStream out = new FileOutputStream(input);
	    
		wb.write(out);
		wb.close();

		out.close();
		in.close();
	}
}
