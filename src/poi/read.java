package poi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Calendar;
import java.util.List;
import java.util.Scanner;

import javax.swing.JFileChooser;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.encryption.InvalidPasswordException;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.hslf.extractor.PowerPointExtractor;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.sl.extractor.SlideShowExtractor;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

public class read 
	{
	static int count = 0;
	public static void main (String [] args) throws InvalidFormatException, IOException
		{
		
		JFileChooser chooser = new JFileChooser();
		chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
		chooser.showOpenDialog(null);
	
		File myFile = chooser.getSelectedFile();
		
		File[] Files = myFile.listFiles();

		Scanner scan = new Scanner(System.in);
		System.out.println("Enter Keyword to Search for");
	
		String keyword = scan.next();
		scan.close();
		
		Calendar cal = Calendar.getInstance();
		
		search(Files,keyword);
		
		System.out.println("\nTotal Items Found: " + count);
		
		System.out.println("\nSTARTED: " + cal.get(Calendar.HOUR) + ":" + cal.get(Calendar.MINUTE) + ":" + cal.get(Calendar.SECOND));
	    cal = Calendar.getInstance();
		System.out.println("FINISHED: "  + cal.get(Calendar.HOUR) + ":" + cal.get(Calendar.MINUTE) + ":" + cal.get(Calendar.SECOND));
	}
	
	public static void search (File [] Files,String keyword)  throws FileNotFoundException   {
		
		try {
		for (int i = 0; i < Files.length; i++) {
		
			if(Files[i].getName().contains(".docx")  && (ReadDocx(Files[i],keyword) == true)) {
				System.out.println("File: " + Files[i].getName());
				count++;
			}
			/*
			if(Files[i].getName().contains(".ppt")  && (ReadPPTX(Files[i],keyword) == true)) {
				System.out.println("File: " + Files[i].getName());		
				count++;
			}*/
			
			if(Files[i].getName().contains(".pdf")  && (ReadPDF(Files[i],keyword) == true)) {
				System.out.println("File: " + Files[i].getName());		
				count++;
			}
			
			if (Files[i].isFile() && Files[i].canRead()) {
				Scanner s = new Scanner(Files[i]);	
				while(s.hasNext()) {
					String data = s.nextLine();
					if(data.toLowerCase().contains(keyword)) {
						System.out.println("File: " + Files[i].getName());
						//System.out.println("FOUND: " + data);
						s.close();
						//Files[i].delete();
						count++;
						break;
				}					
			  }
			}
				if (Files[i].isDirectory() && Files[i].canRead()) {
			    System.out.println("Directory " + Files[i].getName());
			    File NewFile = Files[i];
			    File [] NewList = NewFile.listFiles();
			    search(NewList,keyword);
				}
		}
		
		}catch(NullPointerException e) {
			//e.printStackTrace();
		}catch(Exception e) {
			//e.printStackTrace();
		}
		
	}
	public static boolean ReadDocx (File Files, String keyword) throws InvalidFormatException, IOException 
	{
		XWPFDocument doc = new XWPFDocument(OPCPackage.open(Files));
		@SuppressWarnings("resource")
		XWPFWordExtractor ext = new XWPFWordExtractor(doc);
		String str = ext.getText();
		doc.close();
		
		if(str.toLowerCase().contains(keyword)) {
			return true;
		}
		else {
			return false;
		}

	}
	
	
	public static boolean ReadPPTX (File Files, String keyword) throws InvalidFormatException, IOException 
	{
		XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(Files));
		
		@SuppressWarnings("deprecation")
		PowerPointExtractor extractor = new PowerPointExtractor(new FileInputStream(Files));
		@SuppressWarnings("deprecation")
		String str = extractor.getText(true, true);
		
		/*PowerpointExtractor extract = new PowerpointExtractorppt);
		
		String str = pptText.toString();
		ppt.close();
		pptText.close();
		*/
		if(str.toLowerCase().contains(keyword.toLowerCase())) {
			return true;
		}
		else {
			return false;
		}
	}
	
	public static boolean ReadPDF (File Files, String keyword) throws InvalidPasswordException, IOException {
		PDDocument document = PDDocument.load(Files);
		
		PDFTextStripper Tstripper = new PDFTextStripper();
		
		String str = Tstripper.getText(document);
		
		if(str.toLowerCase().contains(keyword)) {
			return true;
		}
		else {
			return false;
		}
	
	}
}
		
