package poi;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Calendar;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class convert{

	static String outFn = "C:\\Users\\" + getShortname() + "\\Documents\\ChargeBacks\\ExcelReports\\";
    
	public static String getShortname() 
    {
	    String s = System.getProperty("user.name");
	    return s;   
    }
	
    public static File xlsx2xls(File inpFn, Boolean delete) throws InvalidFormatException,IOException 
    {
        InputStream in = new FileInputStream(inpFn);
        
        Calendar cal = Calendar.getInstance();
        
        File outF;
     
        try {
          
			XSSFWorkbook wbIn = new XSSFWorkbook(in);
            new File("C:\\Users\\" + getShortname() + "\\Documents\\ChargeBacks\\ExcelReports").mkdirs();
             outF = new File(outFn + "Converted Excel " + ((cal.get(Calendar.MONTH)+1)) + "-" + (cal.get(Calendar.DAY_OF_MONTH))+ "-" + (cal.get(Calendar.YEAR)) + " " + 
     				(cal.get(Calendar.HOUR)) + "-" + (cal.get(Calendar.MINUTE)) + "-" + (cal.get(Calendar.SECOND)) +".xls");

            
			Workbook wbOut = new HSSFWorkbook();
            int sheetCnt = wbIn.getNumberOfSheets();
            for (int i = 0; i < sheetCnt; i++) {
                Sheet sIn = wbIn.getSheetAt(i);
                Sheet sOut = wbOut.createSheet(wbIn.getSheetName(i));
                Iterator<Row> rowIt = sIn.rowIterator();
                while (rowIt.hasNext()) {
                    Row rowIn = rowIt.next();
                    Row rowOut = sOut.createRow(rowIn.getRowNum());

                    Iterator<Cell> cellIt = rowIn.cellIterator();
                    while (cellIt.hasNext()) {
                        Cell cellIn = cellIt.next();
                        Cell cellOut = rowOut.createCell(cellIn.getColumnIndex(), cellIn.getCellType());

                        switch (cellIn.getCellType()) {
                        case BLANK: break;

                        case BOOLEAN:
                            cellOut.setCellValue(cellIn.getBooleanCellValue());
                            break;

                        case ERROR:
                            cellOut.setCellValue(cellIn.getErrorCellValue());
                            break; 

                        case FORMULA:
                            cellOut.setCellFormula(cellIn.getCellFormula());
                            break;

                        case NUMERIC:
                            cellOut.setCellValue(cellIn.getNumericCellValue());
                            break;

                        case STRING:
                            cellOut.setCellValue(cellIn.getStringCellValue());
                            break;
						default:
							break;
                        }
     
                        {
                         CellStyle styleIn = cellIn.getCellStyle();
                         CellStyle styleOut = cellOut.getCellStyle();
                         styleOut.setDataFormat(styleIn.getDataFormat());
                        }
                        cellOut.setCellComment(cellIn.getCellComment());
                    }
                }
                wbOut.close();
                wbIn.close(); 
            }
            OutputStream out = new BufferedOutputStream(new FileOutputStream(outF));
            try {
                wbOut.write(out);
            } finally {
                out.close();
            }
        } finally {
            in.close();
        }
        if(delete == true) {
        outF.deleteOnExit();}
        return outF;
    }
    
    public static File xls2xlsx (File inpFn, String outdest ) throws InvalidFormatException, IOException 
    {
    	File outF;
    	InputStream in = new BufferedInputStream(new FileInputStream(inpFn));
    	try {

			Workbook wbIn = new HSSFWorkbook(in);
    		outF = new File(outdest);
    		if (outF.exists())
    		outF.delete();
    
    Workbook wbOut = new XSSFWorkbook();
    int sheetCnt = wbIn.getNumberOfSheets();
    for (int i = 0; i < sheetCnt; i++) {
        Sheet sIn = wbIn.getSheetAt(i);
        Sheet sOut = wbOut.createSheet(sIn.getSheetName());
        Iterator<Row> rowIt = sIn.rowIterator();
        while (rowIt.hasNext()) {
        	
            Row rowIn = rowIt.next();
            Row rowOut = sOut.createRow(rowIn.getRowNum());

            Iterator<Cell> cellIt = rowIn.cellIterator();
            while (cellIt.hasNext()) {
                Cell cellIn = cellIt.next();
                Cell cellOut = rowOut.createCell(cellIn.getColumnIndex(), cellIn.getCellType());
                		
                switch (cellIn.getCellType()) {
                case BLANK:
                    break;

                case BOOLEAN:
                    cellOut.setCellValue(cellIn.getBooleanCellValue());
                    break;

                case ERROR:
                    cellOut.setCellValue(cellIn.getErrorCellValue());
                    break;

                case FORMULA:
                    cellOut.setCellFormula(cellIn.getCellFormula());
                    break;

                case NUMERIC:
                    cellOut.setCellValue(cellIn.getNumericCellValue());
                    break;

                case STRING:
                    cellOut.setCellValue(cellIn.getStringCellValue());
                    break;
				default:
					break;
                }
                
                {
                    CellStyle styleIn = cellIn.getCellStyle();
                    CellStyle styleOut = cellOut.getCellStyle();
                    styleOut.setDataFormat(styleIn.getDataFormat());
                }
                cellOut.setCellComment(cellIn.getCellComment());
                wbIn.close();
            }
        }
    }

    OutputStream out = new BufferedOutputStream(new FileOutputStream(outF));
    try {
    	
        wbOut.write(out);
        wbOut.close();
    	} finally {
        out.close();
    		}
    	} finally {
    		in.close();
    	}
    	return outF;
    }
}