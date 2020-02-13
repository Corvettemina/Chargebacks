package poi;

import java.io.File;
import java.io.IOException;

import javax.swing.JFileChooser;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

public class readpdf {
	
	static String BOOK1 = ("C:\\Users\\hannami\\Documents\\test.xls");
	
	public static void main (String [] args) throws IOException, BiffException{
	
        Workbook Book1 = Workbook.getWorkbook(new File(BOOK1));    
        Sheet sheet1 = Book1.getSheet(0);
        
        String [] ReferenceNum = new String [sheet1.getRows()];
        
        for (int k = 0; k < sheet1.getRows(); k++) {
            Cell serial = sheet1.getCell(0, k);
            ReferenceNum[k] = serial.getContents();
            }
        Book1.close();
        
		for (int k = 0; k < ReferenceNum.length; k++) {	
			System.out.println(ReferenceNum[k]);
		}
		
	JFileChooser chooser = new JFileChooser();
	chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
	chooser.showOpenDialog(null);
	
	File myFile = chooser.getSelectedFile();
	
	File[] Files = myFile.listFiles();

	for(int i = 0; i < Files.length; i++) {
		if(Files[i].getName().endsWith(".pdf")&& Files[i].isFile()) {
		PDDocument document = PDDocument.load(Files[i]);
		PDFTextStripper pdfStripper  = new PDFTextStripper();
	    String str = pdfStripper.getText(document);
			
		for (int k = 0; k < ReferenceNum.length; k++) {		
			if(str.contains(ReferenceNum[k])) {
					System.out.println(Files[i].getName());
					break;
						}
					}
				}
			}
	
	}
	}
	
