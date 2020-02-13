package poi;

import java.awt.geom.Rectangle2D;
import java.io.IOException;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.encryption.InvalidPasswordException;
import org.apache.pdfbox.text.PDFTextStripperByArea;

public class Credits {
	
	public static float getCredits () throws InvalidPasswordException, IOException {
		PDDocument document = PDDocument.load(ChargeBacks.statementfile);
		PDPage page = document.getPage(0);
	    										 
	    Rectangle2D Credit = new Rectangle2D.Double(400,670,300,10);
	    String regionName = "credit";
 
	    PDFTextStripperByArea strip = new PDFTextStripperByArea();
	    strip.setSortByPosition(true);
	    strip.addRegion(regionName, Credit);
	    strip.extractRegions(page);
	    
	    String cred = strip.getTextForRegion(regionName);
	    cred = cred.replace("$", "");
	    
	    return Float.valueOf(cred);  
	}
	public static float previousBal () throws InvalidPasswordException, IOException {
		
		PDDocument document = PDDocument.load(ChargeBacks.statementfile);
		PDPage page = document.getPage(0);
	    										 
	    Rectangle2D Credit = new Rectangle2D.Double(400,600,300,10);
	    String regionName = "credit";
 
	    PDFTextStripperByArea strip = new PDFTextStripperByArea();
	    strip.setSortByPosition(true);
	    strip.addRegion(regionName, Credit);
	    strip.extractRegions(page);
	    
	    String cred = strip.getTextForRegion(regionName);
	    
	    if(cred.contains("CR")) 
	    {
	    	cred = cred.replaceAll("[$,CR,CE]", "");
	    	return Float.valueOf(cred);
	    	}
	    else
	    	{
	    	return 0;
	    	}
	    }
}
