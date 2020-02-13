package poi;

import java.io.File;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

public class dropboxOnlineTest {
	public static void main (String [] args) throws InvalidFormatException, IOException{
		
		
		File file = new File("C:\\Users\\minah\\Dropbox\\ARCHIVED Liturgy PPTs\\1-Thoout\\1st Sunday of Thoout (6) 2018.pptx");
		if(file.exists()){
			System.out.println("Exsits");
			
		}
		if(file.canRead()){
			System.out.println("Read");
		}
		if(file.canWrite()){
			System.out.println("Write");
		}
		
		//System.out.println(read.ReadPPTX(file, "Father"));
		
	}

}
