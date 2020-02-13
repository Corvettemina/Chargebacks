package poi;

import java.awt.geom.Rectangle2D;
import java.io.IOException;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.text.PDFTextStripperByArea;

class SelfGenInfo {

	public String getSubtotal () throws IOException 
	
	{
		PDDocument document = PDDocument.load(ChargeBacks.statementfile);
		PDPage page = document.getPage(0);
	    /********************************************x , y , w , h*/
	    Rectangle2D region = new Rectangle2D.Double(400,625,300,10);
	    String regionName = "subtotal";
 
	    PDFTextStripperByArea strip = new PDFTextStripperByArea();
	    strip.setSortByPosition(true);
	    strip.addRegion(regionName, region);
	    strip.extractRegions(page);
	    
	    String _subtotal = strip.getTextForRegion(regionName);

	    _subtotal =  _subtotal.replaceAll("[$, ]" , "");

	    float credits = Credits.getCredits();
	    float prevBal = Credits.previousBal();
	    float subfloat = Float.valueOf(_subtotal);
	    
	    if(credits > 0.00) {
	    	
	    	subfloat = subfloat - credits;
	    	_subtotal = Float.toString(subfloat);
	    }
	    if(prevBal > 0.00) {
	    	subfloat = subfloat - prevBal;
	    	_subtotal = Float.toString(subfloat);
	    }
	    
	    return ("$" + _subtotal);

	}
	public String getStatementDate() throws IOException
	{
		PDDocument document = PDDocument.load(ChargeBacks.statementfile);
		PDPage page = document.getPage(0);
		PDFTextStripperByArea strip = new PDFTextStripperByArea();
		
		Rectangle2D region1 = new Rectangle2D.Double(240,630,100,10);
		strip.addRegion("statementdate", region1);
		strip.extractRegions(page);
		String date = strip.getTextForRegion("statementdate");
		
	    date = date.replace("/", "");

		return date;
	}
	public String getStatementDateWSlash() throws IOException
	{
		PDDocument document = PDDocument.load(ChargeBacks.statementfile);
		PDPage page = document.getPage(0);
		PDFTextStripperByArea strip = new PDFTextStripperByArea();
		
		Rectangle2D region1 = new Rectangle2D.Double(240,630,100,10);
		strip.addRegion("statementdate", region1);
		strip.extractRegions(page);
		String date = strip.getTextForRegion("statementdate");
		
		StringBuffer str = new StringBuffer(date);
		str.insert(6, "20");

		return str.toString();
	}
}