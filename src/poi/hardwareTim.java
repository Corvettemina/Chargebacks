/**
 * *
 * Compares two rows & highlights * */

package poi;

import java.awt.Desktop;
import java.io.File;
import java.io.IOException;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableCell;
import jxl.write.WritableCellFormat;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class hardwareTim {

     static String BOOK1 = ("C:\\users\\hannami\\documents\\tim\\WIN10machines.xls");
     static String BOOK2 = ("C:\\users\\hannami\\documents\\tim\\Hardware.xls");
     
     static String [] WIN10ConsoleUser;
     static String [] HardwareConsoleUser;
     static String [] FirstName;
     static String [] LastName;
     static String [] WIN10ConsoleUserS;
     
    public static void main (String[] args) throws BiffException, IOException, WriteException, RowsExceededException {

        Desktop desktop = Desktop.getDesktop();
             
        Workbook Book1 = Workbook.getWorkbook(new File(BOOK1));
        WritableWorkbook ww = Workbook.createWorkbook(new File(BOOK1),Book1);       
        Sheet Win10Sheet = Book1.getSheet(0);
        //WritableSheet Win10WRITABLE = ww.getSheet(0);
        
        Workbook Book2 = Workbook.getWorkbook(new File(BOOK2));        
        WritableWorkbook ww2 = Workbook.createWorkbook(new File(BOOK2),Book2);
        Sheet HARDWARE = Book2.getSheet(0);
        WritableSheet HardwareWRITEABLE = ww2.getSheet(0);
        
        int max = 0;
        if(Win10Sheet.getRows() < HARDWARE.getRows()) {
        	max = HARDWARE.getRows();
        }
        
        if(Win10Sheet.getRows() > HARDWARE.getRows()) {
        	max = Win10Sheet.getRows();
        }
        
        if(Win10Sheet.getRows() == HARDWARE.getRows()) {
        	max = Win10Sheet.getRows();
        }
        
        WIN10ConsoleUser = new String[max];
        HardwareConsoleUser = new String[max];
        FirstName = new String[max];
        LastName = new String[max];
        WIN10ConsoleUserS = new String [max];     

       
        for (int i = 0; i < Win10Sheet.getRows(); i++) {
            Cell cell = Win10Sheet.getCell(3, i);
            WIN10ConsoleUser[i] = cell.getContents();
        	}  
            
         for (int i = 0; i < Win10Sheet.getRows(); i++) {
            Cell cell = Win10Sheet.getCell(0, i);
            WIN10ConsoleUserS[i] = cell.getContents();
        	}
        
        for (int k = 0; k < HARDWARE.getRows(); k++) {
            Cell cell1 = HARDWARE.getCell(0, k);
            HardwareConsoleUser[k] = cell1.getContents();
            }
        for (int k = 0; k < Win10Sheet.getRows(); k++) {
        	Cell first = Win10Sheet.getCell(1, k);
        	FirstName[k] = first.getContents(); 
        }
        
        for (int k = 0; k < Win10Sheet.getRows(); k++) {
        	Cell last = Win10Sheet.getCell(2, k);
        	LastName[k] = last.getContents(); 
        }
        
        
        if(Win10Sheet.getRows() < HARDWARE.getRows()) 
        {
        	for( int i = Win10Sheet.getRows(); i < HARDWARE.getRows(); i++) {
        	WIN10ConsoleUser [i] = ""; 
  	     	}
  	    }
        
        if(HARDWARE.getRows() < Win10Sheet.getRows()) 
        {
        	for( int i = HARDWARE.getRows(); i < Win10Sheet.getRows(); i++) {
        	HardwareConsoleUser[i] = "";
    	    }
  	    }
 
           Compare(HardwareWRITEABLE); 
 
           ww.write();
           ww.close();
           Book1.close();
           
           ww2.write();
           ww2.close();
           Book2.close();
           
           //desktop.open(new File(BOOK2));
           desktop.open(new File(BOOK2));
           
          System.exit(0);
        }
        
       public static void Compare(WritableSheet HardwareWRITEABLE) throws WriteException {

        
        for (int i = 0; i < WIN10ConsoleUser.length; i++) {
            for (int j = 0; j < HardwareConsoleUser.length; j++) {
                if ((WIN10ConsoleUser[i].equalsIgnoreCase(HardwareConsoleUser[i])) || (WIN10ConsoleUser[i].equalsIgnoreCase(HardwareConsoleUser[j]))) {
                	   
                     Label FNameLabel = new Label (4,j,FirstName[i]);
                	 HardwareWRITEABLE.addCell(FNameLabel);
                	
                  	 Label LNameLabel = new Label (5,j,LastName[i]);
                  	 HardwareWRITEABLE.addCell(LNameLabel);
                     
                     /* Label Shortname = new Label (3,j,WIN10ConsoleUserS[j]);
                  	 HardwareWRITEABLE.addCell(Shortname); */
                    
                     WritableCell c = HardwareWRITEABLE.getWritableCell(3,j);
                     WritableCellFormat newFormat = new WritableCellFormat();
                     //newFormat.setBackground(Color.YELLOW);
                     c.setCellFormat(newFormat);
                }
            }
        } 
        /***   
       List<String> myList = new ArrayList<String>(com);
       List<String> myList1 = new ArrayList<String>(FName);
       List<String> myList2 = new ArrayList<String>(LName);
      
       for(int i = 0; i < myList.size(); i ++)	
       		{
    	   		if(myList.get(i).equals("")){
    	   		myList.remove(i);}
       		}
       
       System.out.println(myList.toString());
       System.out.println();
       System.out.println("Total Elements in Common: " + myList.size());
       
       /*for(int i = 0; i < myList.size(); i ++){
       Label label4 = new Label (5,i,myList.get(i));
       ws.addCell(label4);
       } 
       
       for(int i = 0; i < myList1.size(); i ++){
           Label label4 = new Label (1,(i+1),myList1.get(i));
           Win10WRITABLE.addCell(label4);
           } 
       
       for(int i = 0; i < myList2.size(); i ++){
           Label label4 = new Label (2,(i+1),myList2.get(i));
           Win10WRITABLE.addCell(label4);
           } */
    }

}