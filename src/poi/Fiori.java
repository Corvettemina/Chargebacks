package poi;

import java.awt.Desktop;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Calendar;

import javax.swing.JFileChooser;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.UIManager;
import javax.swing.UnsupportedLookAndFeelException;
import javax.swing.filechooser.FileNameExtensionFilter;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

public class Fiori extends Format  {

	static ArrayList<String> costExcel, costCenter, io, Item, quantity;
	static Workbook chargebacks;

	public static void main () throws BiffException, IOException, WriteException, UnsupportedLookAndFeelException,
										 ClassNotFoundException, InstantiationException, IllegalAccessException, InvalidFormatException {

		UIManager.setLookAndFeel(ChargeBacks.lookAndFeel);

		Calendar cal = Calendar.getInstance();

		JFileChooser chooser = new JFileChooser();
		FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel Files", "xlsx","xls");
		chooser.setFileFilter(filter);
	    chooser.setDialogTitle("Open Excel Sheet");
		chooser.showOpenDialog(null);
		File book = chooser.getSelectedFile();

		if(book.getName().endsWith(".xlsx")) {
			book = xlsx2xls(book, true);
		}

		chargebacks = Workbook.getWorkbook(book);
		Sheet ChargebacksSheet = chargebacks.getSheet(0);
		WritableWorkbook newBook = Workbook.createWorkbook(book,chargebacks);

		WritableSheet ChargebackWS = newBook.getSheet(0);
		WritableSheet newBookWrite = newBook.createSheet("FIORI on " + ((cal.get(Calendar.MONTH)+1)) +
				"-" + (cal.get(Calendar.DAY_OF_MONTH))+ "-" + (cal.get(Calendar.YEAR)) + " " + (cal.get(Calendar.HOUR)) + "-" +
				(cal.get(Calendar.MINUTE)) + "-" + (cal.get(Calendar.SECOND))  ,chargebacks.getNumberOfSheets()+1);


		costExcel = new ArrayList<String>();
		costCenter = new ArrayList<String>();
		io = new ArrayList<String>();
		Item = new ArrayList<String>();
		quantity = new ArrayList<String>();

		int max = ChargebacksSheet.getRows();

		for(int i = 0; i < max; i++)
		 {
			if((ChargebacksSheet.getCell(10, i).getContents().equalsIgnoreCase("open") || ChargebacksSheet.getCell(10, i).getContents().equalsIgnoreCase("")) &&
					!ChargebacksSheet.getCell(2, i).getContents().equalsIgnoreCase("10900751"))
			{
			costCenter.add(ChargebacksSheet.getCell(2, i).getContents());
			io.add(ChargebacksSheet.getCell(3, i).getContents());
			Item.add(ChargebacksSheet.getCell(5, i).getContents());
			costExcel.add(ChargebacksSheet.getCell(6, i).getContents());
			quantity.add(ChargebacksSheet.getCell(7, i).getContents());
			ChargebackWS.addCell(new Label(10,i,"Pending"));
			}
		 }

		chargebacks.close();

	new File("C://users//" + convert.getShortname() + "//documents//Chargebacks//FIORI").mkdirs();

    String [] field= {"Record type\nHeader = 1\nDetails = 2" ,"Company\nCode","Approving\nCost Center", "Description" , "Posting Key" ,
    							"Account", "Amount", "Cost Center" ,"Internal order", "Profit Center", "Request note"};

	String [] NoOfChar = {"1","US01","10900751"};

	float totalcost = 0;
	for(int i = 1; i<quantity.size(); i++) {
		if(!quantity.get(i).equals("1") && !quantity.get(i).equals("")) {
			totalcost = (Float.valueOf(quantity.get(i)) * Float.valueOf(costExcel.get(i)));
			costExcel.set(i , Float.toString(totalcost));
		}
	}

	for(int i = 0; i< field.length; i++) {
		newBookWrite.addCell(new Label(i,0,field[i]));
	if(i <  NoOfChar.length) {
		newBookWrite.addCell(new Label(i,1,NoOfChar[i]));
		}
	}
		newBookWrite.addCell(new Label(3,1,JOptionPane.showInputDialog("Enter Description")));

		float total = 0;

		for(int i = 1; i < costExcel.size(); i++) {
			if(!costExcel.get(i).equals("")) {
				total = total + Float.valueOf(costExcel.get(i));
			}
		}

		newBookWrite.addCell(new Label(0,2,"2"));
		newBookWrite.addCell(new Label(4,2,"50"));
		newBookWrite.addCell(new Label(5,2,"614090"));
		newBookWrite.addCell(new Label(6,2,Double.toString(total)));
		newBookWrite.addCell(new Label(7,2,"10900751"));
		newBookWrite.addCell(new Label(8,2,"7US010000176"));
		newBookWrite.addCell(new Label(10,2,"Charges for Ghost Card"));

		for(int i = 0; i < Item.size(); i++) {		if(Item.get(i).length() > 50) {

				JPanel panel = new JPanel();
				panel.setToolTipText("Input Over 50 Characters. Input a new Item Description");
				JLabel label = new JLabel(Item.get(i));

				panel.add(label);
				Item.set(i , JOptionPane.showInputDialog(panel, Item.get(i)));
		}
	}

	for(int i = 0; i < costCenter.size(); i++)
	{
			newBookWrite.addCell(new Label(0,i+3,"2"));
			newBookWrite.addCell(new Label(4,i+3,"40"));
			newBookWrite.addCell(new Label(5,i+3,"614090"));
			newBookWrite.addCell(new Label(6,i+3,costExcel.get(i)));
			newBookWrite.addCell(new Label(7,i+3,costCenter.get(i)));
			newBookWrite.addCell(new Label(8,i+3,io.get(i)));
			newBookWrite.addCell(new Label(10,i+3,Item.get(i)));
	}

	newBook.write();
	newBook.close();

	if(chooser.getSelectedFile().getName().endsWith(".xlsx")) {
		book = xls2xlsx(book,chooser.getSelectedFile().getPath());
	}

	SizeCols(book);

	Desktop d = Desktop.getDesktop();
	d.open(book);

	}
}
