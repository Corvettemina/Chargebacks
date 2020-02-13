package poi;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;

import javax.swing.JFileChooser;
import javax.swing.filechooser.FileNameExtensionFilter;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;

public class readDocx {
	public static void main (String [] args) throws Exception {
		
		JFileChooser chooser = new JFileChooser();
		FileNameExtensionFilter filter = new FileNameExtensionFilter("Word Document Files", "doc");   
		chooser.setFileFilter(filter);      
		chooser.showOpenDialog(null); 
		
		File file = chooser.getSelectedFile();
		InputStream input = new FileInputStream(file);
		
		HWPFDocument doc = new HWPFDocument(input);
		WordExtractor ext = new WordExtractor(doc);
		String str = ext.getText();
		doc.close();
		ext.close();
		
		if(str.toLowerCase().contains("mina")) {
		System.out.println(str);}
		
	}
}
