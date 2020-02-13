package poi;

import java.awt.Desktop;
import java.io.File;
import java.io.FileNotFoundException;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Iterator;
import java.util.List;

import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.UIManager;
import javax.swing.filechooser.FileNameExtensionFilter;

import org.apache.pdfbox.io.MemoryUsageSetting;
import org.apache.pdfbox.multipdf.PDFMergerUtility;
import org.apache.pdfbox.multipdf.Splitter;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

public class ChargeBacks {
	
	 static String lookAndFeel = "com.sun.java.swing.plaf.windows.WindowsLookAndFeel";
     static String shortname = "";
     static File statementfile = null;
	
	public static void main () throws  Exception{
		   
		UIManager.setLookAndFeel(lookAndFeel);
		
		JFileChooser chooser = new JFileChooser();
		FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel Files", "xlsx", "xls");
		chooser.setFileFilter(filter);
        chooser.setDialogTitle("Open 'USBank Transactions' Excel Sheet");
		chooser.showOpenDialog(null);
	    File USBankData = chooser.getSelectedFile();
	    chooser.setCurrentDirectory(USBankData);
	    
	    shortname = convert.getShortname();
	   
	    if(USBankData.getName().endsWith(".xlsx")) {
	    USBankData = convert.xlsx2xls(USBankData,true);
	    }
	    
	    Application.progressBar.setValue(10);
	    Application.lblNewLabel.setText("Progress: 10%");
	    String [] PurchaseID = null;
	    String [] Cost = null;
	    int purchaseidnum = 0;
	    int costidnum = 0;
	   
		try {
	    Workbook Book = Workbook.getWorkbook(USBankData);    
        Sheet sheet = Book.getSheet(0); 
        
        for(int i = 0; i < sheet.getColumns(); i++) {
        if(sheet.getCell(i,0).getContents().equals("Purchase ID")) {
        	purchaseidnum = i;
        	}	
        if(sheet.getCell(i,0).getContents().equals("Source Currency Amount")) {
        	costidnum = i; 
        	}
        		if(purchaseidnum > 0 && costidnum > 0) {
        			break;
        		}
        }

        PurchaseID = new String [sheet.getRows()];
        Cost = new String [sheet.getRows()];

        Application.progressBar.setValue(15);
	    Application.lblNewLabel.setText("Progress: 15%");

	    for (int k = 0; k < sheet.getRows(); k++) {
            PurchaseID[k] = sheet.getCell(purchaseidnum, k).getContents();

            Cost[k] = sheet.getCell(costidnum, k).getContents();
            if(!Cost[k].equals("Source Currency Amount")) {
            if(Float.valueOf(Cost[k]) > 999.99 && Float.valueOf(Cost[k]) <= 9999.99 ) {
            	Cost[k] = Cost[k].substring(0,1) + "," + Cost[k].substring(1);
            	}
            }
        }

        Book.close();
        }catch(BiffException e){
        	System.err.println(e);
        }

		for(int i = 0; i < PurchaseID.length; i++) {
		    if(PurchaseID[i].equals("")) {
		    	PurchaseID[i] = "NO UNUSED PURCHASE IDS";
		    	}
		    if(Cost[i].equals("")) {
		    	Cost[i] = "NO UNUSED PURCHASE IDS";
		    	}
	    }

		Application.progressBar.setValue(20);
		Application.lblNewLabel.setText("Progress: 20%");
		
		filter = new FileNameExtensionFilter("PDF Files", "pdf");
        chooser.setFileFilter(filter);
        chooser.setDialogTitle("Open PDF of Recipets");
		chooser.showOpenDialog(null);
	    File file = chooser.getSelectedFile();
	    chooser.setCurrentDirectory(file);

	    Application.progressBar.setValue(30);
	    Application.lblNewLabel.setText("Progress: 30%");
	    
        chooser.setDialogTitle("Open Statment File");
		chooser.showOpenDialog(null);
	    statementfile = chooser.getSelectedFile();

	    Application.progressBar.setValue(40);
        Application.lblNewLabel.setText("Progress: 40%");

		//Opens PDF 
	    PDDocument OrginalFile = PDDocument.load(file);

	    Splitter split = new Splitter();
		List<PDDocument> Pages = split.split(OrginalFile); 	 
        Iterator<PDDocument> iterator = Pages.listIterator(); 
 
        new File("C://users//" + shortname + "//documents//chargebacks//temp").mkdirs();
        new File("C://users//" + shortname + "//documents//chargebacks//receipts").mkdir();
        
        Application.progressBar.setValue(40);
        Application.lblNewLabel.setText("Progress: 40%");
        
        int z = 0;
        
        while(iterator.hasNext()) { 
            PDDocument pd = iterator.next(); 
            pd.save("C://users//" + shortname + "//documents//chargebacks//temp//" + ++z + ".pdf");     
         } 

		File directory = new File("C://users//" + shortname + "//documents//chargebacks//temp");
		File [] myFiles = directory.listFiles();
		
		List<String> Common = new ArrayList<String>();
		
		Application.progressBar.setValue(50);
		Application.lblNewLabel.setText("Progress: 50%");
		
		for(int i = 0; i<myFiles.length; i++){
				
				if(myFiles[i].getName().endsWith(".pdf")){
				PDDocument document = PDDocument.load(myFiles[i]);
				PDFTextStripper textstrip = new  PDFTextStripper();
				String st = textstrip.getText(document).toLowerCase();
			
				int count = 0;
				for(int k = 0; k< PurchaseID.length; k++)
					{
					if(st.contains(PurchaseID[k]) && st.contains("$ "+ Cost[k])) {
				    
					Common.add(PurchaseID[k]);
					document.close();	
					break;
					}else {
						count++;
					}
					if(count == PurchaseID.length) {	
						document.close();
						myFiles[i].delete();
					}
				  }
				}
			  }
		
		Application.progressBar.setValue(50);
		Application.lblNewLabel.setText("Progress: 50%");
		
		Calendar cal = Calendar.getInstance();
		String filePath = ("C:\\Users\\" + shortname + "\\Documents\\ChargeBacks\\receipts\\FINAL_" +
				((cal.get(Calendar.MONTH)+1)) + "-" + (cal.get(Calendar.DAY_OF_MONTH))+ "-" + (cal.get(Calendar.YEAR)) + " " + 
				(cal.get(Calendar.HOUR)) + "-" + (cal.get(Calendar.MINUTE)) + "-" + (cal.get(Calendar.SECOND)) +".pdf");
		
		PDFMergerUtility merger = new PDFMergerUtility();
		merger.setDestinationFileName(filePath);
		try {
		merger.addSource(SelfGen.getSelfGen());
		}catch(FileNotFoundException e) {
			SelfGen.end();
		}
		merger.addSource(statementfile);
		
		/****************************
		Merges matched sheets to one*/ 
		
		myFiles = directory.listFiles();
		for(int i = 0; i <myFiles.length; i++) {
			merger.addSource(myFiles[i]);
		}

		merger.mergeDocuments(MemoryUsageSetting.setupMainMemoryOnly());
		
		Application.progressBar.setValue(60);
		Application.lblNewLabel.setText("Progress: 60%");
		
		for(int i = 0; i< myFiles.length; i++) {
			myFiles[i].delete();
		}
		
		Application.progressBar.setValue(70);
		Application.lblNewLabel.setText("Progress: 70%");
		
		new File("C://users//" + shortname + "//documents//chargebacks//temp").delete();
		
		List<String> Ref = new ArrayList<String>();
		for(int k = 0; k< PurchaseID.length; k++){
			Ref.add(PurchaseID[k]);
		}
		
		Common.add("Purchase ID");
		Ref.removeAll(Common);
		
		
		Runtime.getRuntime().exec("attrib +h /s /d" ,null, new File("O:/4sam"));
		
		Application.progressBar.setValue(80);
		Application.lblNewLabel.setText("Progress: 80%");
		
		if(Ref.size()>1) 
		{
			Application.progressBar.setValue(90);
			Application.lblNewLabel.setText("Progress: 90%");
			Ref.remove("NO UNUSED PURCHASE IDS");
			JOptionPane.showMessageDialog(Application.contentPane,  Ref.size() + " Unused Purchase IDs: " + Ref);
		    System.out.println(Ref);
			Desktop.getDesktop().open(Unused.WriteToFile(Ref));
		}else{
			Application.progressBar.setValue(90);
			Application.lblNewLabel.setText("Progress: 90%");
			JOptionPane.showMessageDialog(Application.contentPane,  "0 Unused Purchase IDs");
		}	
		
		Application.progressBar.setValue(100);
		Application.lblNewLabel.setText("Progress: 100%");

		Desktop.getDesktop().open(new File(filePath));	
		
		}	
	}