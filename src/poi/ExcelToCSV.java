package poi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import javax.swing.JFileChooser;
import javax.swing.filechooser.FileNameExtensionFilter;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelToCSV {

public static void convertToXlsx(File inputFile, File outputFile) 
{
	// For storing data into CSV files
StringBuffer cellValue = new StringBuffer();
try 
{
	FileOutputStream fos = new FileOutputStream(outputFile);

	// Get the workbook instance for XLSX file
	XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(inputFile));

	// Get first sheet from the workbook
	XSSFSheet sheet = wb.getSheetAt(0);

	Row row;
	Cell cell;

	// Iterate through each rows from first sheet
	Iterator<Row> rowIterator = sheet.iterator();

	while (rowIterator.hasNext()) 
	{
	row = rowIterator.next();

	// For each row, iterate through each columns
	Iterator<Cell> cellIterator = row.cellIterator();
	
	while (cellIterator.hasNext()) 
	{
		cell = cellIterator.next();

		switch (cell.getCellType()) 
		{
		
		case BOOLEAN:
			cellValue.append(cell.getBooleanCellValue() + ",");
			break;
		
		case NUMERIC:
			cellValue.append(cell.getNumericCellValue() + ",");
			break;
		
		case STRING:
			cellValue.append(cell.getStringCellValue() + ",");
			break;

		case BLANK:
			cellValue.append("" + ",");
			break;
			
		default:
			cellValue.append(cell + ",");

		}
	}
	}

fos.write(cellValue.toString().getBytes());
wb.close();
fos.close();

} 
catch (Exception e) 
{
	System.err.println("Exception :" + e.getMessage());
}
}

public static void convertToXls(File inputFile, File outputFile) 
{
// For storing data into CSV files
StringBuffer cellDData = new StringBuffer();
try 
{
	FileOutputStream fos = new FileOutputStream(outputFile);

	// Get the workbook instance for XLS file
	HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(inputFile));
	// Get first sheet from the workbook
	HSSFSheet sheet = workbook.getSheetAt(0);
	Cell cell;
	Row row;

	// Iterate through each rows from first sheet
	Iterator<Row> rowIterator = sheet.iterator();
	while (rowIterator.hasNext()) 
	{
	row = rowIterator.next();

	// For each row, iterate through each columns
	Iterator<Cell> cellIterator = row.cellIterator();
	while (cellIterator.hasNext()) 
	{
	cell = cellIterator.next();

	switch (cell.getCellType()) 
	{
	
	case BOOLEAN:
		cellDData.append(cell.getBooleanCellValue() + ",");
		break;
	
	case NUMERIC:
		cellDData.append(cell.getNumericCellValue() + ",");
		break;
	
	case STRING:
		cellDData.append(cell.getStringCellValue() + ",");
		break;

	case BLANK:
		cellDData.append("" + ",");
		break;
		
	default:
		cellDData.append(cell + ",");
	}
	}
	}

fos.write(cellDData.toString().getBytes());
fos.close();
workbook.close();

}
catch (FileNotFoundException e) 
{
    System.err.println("Exception" + e.getMessage());
} 
catch (IOException e) 
{
	System.err.println("Exception" + e.getMessage());
}
}

public static void main(String[] args) 
{
	JFileChooser chooser = new JFileChooser();
	FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel Files", "xlsx");
	chooser.setFileFilter(filter);
    chooser.setDialogTitle("Open Excel Sheet");
	chooser.showOpenDialog(null);
	File book = chooser.getSelectedFile();
	
	
	
	File inputFile = book;
	File outputFile = new File("C://users//hannami//documents//output1.csv");
	//File inputFile2 = new File("C:\\input.xlsx");
	//File outputFile2 = new File("C:\\output2.csv");
	//convertToXls(inputFile, outputFile);
	convertToXlsx(inputFile, outputFile);
}
}