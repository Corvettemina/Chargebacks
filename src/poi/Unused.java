package poi;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.util.List;

public class Unused {
	public static File WriteToFile (List<String> Ref) throws IOException{
		
	File unused = new File("C://users//" + System.getProperty("user.name") + "//documents//chargebacks//receipts//Unused Purchase IDs.txt");
	unused.createNewFile();
	String write = (Ref.size() + " UNUSED PURCHASE IDS:" + "\r\n");
	for(int i = 0; i < Ref.size(); i++) {
		
		write = write + (Ref.get(i) + "\r\n");
	}
	FileWriter writer = new FileWriter(unused);
	//System.out.println(write);
	writer.write(write);
	writer.flush();
	writer.close();
	
	return unused;
	
	}
}
