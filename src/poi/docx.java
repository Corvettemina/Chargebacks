package poi;

import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

public class docx {
	public static void main (String [] args) throws Exception {
		
		XWPFDocument document = new XWPFDocument(); 

	      FileOutputStream out = new FileOutputStream( new File("C://users//hannami/documents//createdocument.docx"));
	      document.write(out);
	      out.close();
	      document.close();
	      
	      System.out.println("createdocument.docx written successully");

	}

}
