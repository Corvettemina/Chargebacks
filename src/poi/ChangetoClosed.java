package poi;

import java.awt.Desktop;
import java.io.File;
import java.io.IOException;

import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.UIManager;
import javax.swing.UnsupportedLookAndFeelException;
import javax.swing.filechooser.FileNameExtensionFilter;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class ChangetoClosed {
	
	static String [] SheetNames = null;
	static Workbook Fiori;
	static WritableWorkbook wb;
	
	public static void main () throws InvalidFormatException, IOException, BiffException, RowsExceededException, WriteException, 
									ClassNotFoundException, InstantiationException, IllegalAccessException, UnsupportedLookAndFeelException, InterruptedException {
		
		UIManager.setLookAndFeel(ChargeBacks.lookAndFeel);
		
		JFileChooser chooser = new JFileChooser();
		FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel Files", "xlsx","xls");
		chooser.setFileFilter(filter);
	    chooser.setDialogTitle("Open Excel Sheet");
		chooser.showOpenDialog(null);
		File book = chooser.getSelectedFile();
		
		if(book.getName().endsWith(".xlsx")) {
			book = convert.xlsx2xls(book,true);
		}
				
	
		 Fiori = Workbook.getWorkbook(book);
		 wb = Workbook.createWorkbook(new File(book.getAbsolutePath()),Fiori);
		 SheetNames = Fiori.getSheetNames();
		 
		 int nums = 0;
		
		 for(int i = 0; i < SheetNames.length;i++) {
			 String s = JOptionPane.showInputDialog(null, "Close Items in sheet '" + SheetNames[i] + "?' (y/n)'");
			 if(s.equalsIgnoreCase("y")) {
			 nums = i;
			 break;}
		 }
		 
		Change(nums);
		
		if(chooser.getSelectedFile().getName().endsWith(".xlsx")) {
			book = convert.xls2xlsx(book, chooser.getSelectedFile().getPath());
			}

			Format.SizeCols(book);
			
			Desktop d = Desktop.getDesktop();
			d.open(book);
		}
	
	public static void Change(int nums) throws RowsExceededException, WriteException, IOException {
		Sheet sheet = Fiori.getSheet(nums);
	
		WritableSheet ws = wb.getSheet(nums);

		for(int i = 1; i < sheet.getRows(); i ++) {
			if(sheet.getCell(10, i).getContents().equalsIgnoreCase("Pending")) {
				ws.addCell(new Label(10,i,"Closed"));
			}
		}
		
		wb.write();
		Fiori.close();
		wb.close();
	
	}
}
