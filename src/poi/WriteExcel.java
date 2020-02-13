package poi;

import java.awt.Desktop;
import java.io.EOFException;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

import org.apache.poi.EmptyFileException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class WriteExcel {

	public static void main(String[] args) throws IOException, InvalidFormatException, EOFException, EmptyFileException {

		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Sheet1");
		
		XSSFRow row = sheet.createRow(0);
		XSSFCell cell = row.createCell(0);
		cell.setCellValue("1. Cell");

		cell = row.createCell(1);

		DataFormat format = workbook.createDataFormat();

		CellStyle dateStyle = workbook.createCellStyle();
		dateStyle.setDataFormat(format.getFormat("mm.dd.yyyy"));
		cell.setCellStyle(dateStyle);
		cell.setCellValue(new Date());
		row.createCell(2).setCellValue("34. Cell");
		sheet.autoSizeColumn(1);
		workbook.write(new FileOutputStream("c://users//hannami//documents//tim//excel.xlsx"));
      workbook.close();
      
          Desktop desktop = Desktop.getDesktop();          
           desktop.open(new File("c://users//hannami//documents//tim//excel.xlsx"));
           System.exit(0);
}
	
}