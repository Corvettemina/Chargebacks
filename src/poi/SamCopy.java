package poi;


import java.io.File;
import java.io.IOException;
import java.util.Iterator;
import java.util.LinkedHashSet;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class SamCopy {
	public static void main (String [] args) throws InvalidFormatException, IOException{
	/*JFileChooser chooser = new JFileChooser();
	FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel Files", "xlsx", "xls");
	chooser.setFileFilter(filter);
  
	chooser.showOpenDialog(null);*/
    File data = new File("C://users//hannami/documents//Scan_Results_dcx_ak43_20191009_scan_1570460611_59295.xlsx");
    

    LinkedHashSet<String> OS = new LinkedHashSet<String>();
    
    XSSFWorkbook wbIn = new XSSFWorkbook(data);

    //Workbook wbOut = new HSSFWorkbook();
    int sheetCnt = wbIn.getNumberOfSheets();
    for (int i = 0; i < sheetCnt; i++) {
        Sheet sIn = wbIn.getSheetAt(i);
        Sheet sOut = wbIn.createSheet("Copy");
        Iterator<Row> rowIt = sIn.rowIterator();
        while (rowIt.hasNext()) {
            Row rowIn = rowIt.next();
            Row rowOut = sOut.createRow(rowIn.getRowNum());

            Iterator<Cell> cellIt = rowIn.cellIterator();
            /*while (cellIt.hasNext()) {*/
                Cell cellIn = cellIt.next();
                Cell cellOut = rowOut.createCell(cellIn.getColumnIndex(), cellIn.getCellType());

                switch (cellIn.getCellType()) {
                case BLANK: break;

                case STRING:
                    cellOut.setCellValue(cellIn.getStringCellValue());
                   
                    	OS.add(cellIn.getStringCellValue());
				default:
					break;
                    	                      }
                 
				
                }

             
        }
    wbIn.close();

    //}
    //FileOutputStream out = new FileOutputStream(data.getPath());
    
    //wbIn.write(out);
    //wbIn.close();
   // out.close();


    /*ArrayList<String> OS1 = new ArrayList<String>(OS);
    for(int i = 0; i < OS1.size(); i ++) {
    	if(!OS1.get(i).contains("windows 7")) {
    		OS1.remove(i);
    	}
    }*/
    
     System.out.println(OS);
	}
}
