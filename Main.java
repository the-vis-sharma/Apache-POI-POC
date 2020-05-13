import java.io.File;
import java.io.FileOutputStream;
 
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
 
public class Main {
 
	public static void main(String args[]) {
 
		XWPFDocument document = null;
		FileOutputStream fileOutputStream = null;
		try {
 
			document = new XWPFDocument();
			File fileToBeCreated = new File("POI_Table_Test.docx");
			fileOutputStream = new FileOutputStream(fileToBeCreated);
 
			// Create a Simple Table using the document.
			XWPFTable table = document.createTable();
 
			// Now add Rows and Columns to the Table.
			// Creating the First Row
			XWPFTableRow tableRow0 = table.getRow(0);
			
			// Creating the First Cell
			XWPFTableCell tableCell0 = tableRow0.getCell(0);
			tableCell0.setText(" Row 0 Column 0 ");
			
			// Creating the Other Cells for the First Row
			XWPFTableCell tableCell1 = tableRow0.addNewTableCell();
			tableCell1.setText(" Row 0 Column 1 ");
			XWPFTableCell tableCell2 = tableRow0.addNewTableCell();
			tableCell2.setText(" Row 0 Column 2 ");
 
			// Creating the Next Rows and Cells
			XWPFTableRow tableRow1 = table.createRow();
			tableRow1.getCell(0).setText(" Row 1 Column 0 ");
			tableRow1.getCell(1).setText(" Row 1 Column 1 ");
			tableRow1.getCell(2).setText(" Row 1 Column 2 ");
 
			document.write(fileOutputStream);
			fileOutputStream.close();
			System.out.println("Table created in Word File Succefully !!!");
 
		} catch (Exception e) {
			System.out.println("We had an error while creating the Word Doc " + e.getMessage());
			e.printStackTrace();
		}
 
	}
}