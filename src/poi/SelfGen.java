package poi;

import java.io.File;
import java.io.IOException;

import javax.swing.JOptionPane;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDDocumentCatalog;
import org.apache.pdfbox.pdmodel.interactive.form.PDAcroForm;
import org.apache.pdfbox.pdmodel.interactive.form.PDField;



public class SelfGen extends SelfGenInfo
	{
	static File find;
	
	public static void findLocation() 
	{		
		for(int i = 65; i < 91; i++) {	

		find = new File(Character.toString((char)i) + "://4sam//");

		if (!find.exists()){
			find = 	new File(Character.toString((char)(i + 1)) + "://4sam//");

		}
		else {break;}
		}
		/*System.out.println("Found file at " + find.getPath());*/
	}

	public static File getSelfGen () throws IOException

		{
		findLocation();
		File file = new File(find.getPath() +  "//SelfGen - Copy.pdf");
		if(!file.exists()) {
			JOptionPane.showMessageDialog(null, "NO SELF GEN TEMPLATE FILE FOUND\n"
					+ " CHECK 'PCOMMON' DRIVE AND TRY AGAIN");
			end();
		}else {
			Runtime.getRuntime().exec("attrib -h /s /d" ,null, new File(find.getPath()));
			try {
			Thread.sleep(500);
			}catch(Exception e) {}
		}
		
		SelfGenInfo self = new SelfGenInfo();

		PDDocument selfgenPDF = PDDocument.load(file);	
		PDDocumentCatalog docCatalog = selfgenPDF.getDocumentCatalog();
	    PDAcroForm acroForm = docCatalog.getAcroForm();
	    
	    PDField fieldDate = acroForm.getField("Date");
	    fieldDate.setValue(self.getStatementDateWSlash());
	    
	    PDField quantity = acroForm.getField("QUANTITYRow1");
	    quantity.setValue("1");
	    
	    PDField fieldRate = acroForm.getField("RATERow1");
	    String sub = self.getSubtotal();
	    String rate = sub.replace("$","");
	    fieldRate.setValue(rate);
	    
	    PDField [] SubField = {acroForm.getField("Total"), acroForm.getField("AMOUNTRow1"), acroForm.getField("Sub-Total")};
	    
	    for(int i = 0; i < SubField.length; i++) {
	    	SubField[i].setValue(sub);
	    }
	    
	    PDField RefrenceNum = acroForm.getField("Ref #");  
	    RefrenceNum.setValue("8537" +  self.getStatementDate());

	    selfgenPDF.save(file);
	    
	    return file;
	}
	
	public static void end() {
		return;
		}
	
}
