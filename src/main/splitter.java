package main;

// Using AWT containers and components
import java.awt.BorderLayout;
import java.awt.FlowLayout;
import java.awt.GridLayout;
// Using AWT events and listener interfaces
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Iterator;

// Using Swing components and containers
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JTextField;
import javax.swing.SwingConstants;
import javax.swing.SwingUtilities;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class splitter extends JFrame {
	private JTextField  tfColumnId, tfInputFilePath;
	public String fileID;
	public JFileChooser chooser;

	/** Constructor to setup the GUI */
	public splitter() {
		// Retrieve the content-pane of the top-level container JFrame
		// All operations done on the content-pane
		JPanel panelHeader = new JPanel(new FlowLayout());
		panelHeader.setLayout(new BorderLayout());
		panelHeader.add(new JLabel("Configuration Settings", SwingConstants.LEFT));
		JPanel panelConfig = new JPanel(new GridLayout(3, 2, 10, 10));
		panelConfig.add(new JLabel("Enter the column number (eg., 3):"));
		tfColumnId = new JTextField(10);
		panelConfig.add(tfColumnId);
		tfInputFilePath = new JTextField(10);
		panelConfig.add(tfInputFilePath);
		JButton btnBrowse = new JButton("Browse");
		panelConfig.add(btnBrowse);
		JButton btnSend = new JButton("START SPLIT");
		panelConfig.add(btnSend);
		JButton btnReset = new JButton("CLEAR");
		panelConfig.add(btnReset);
		this.setLayout(new BorderLayout());
		panelHeader.add(panelConfig, BorderLayout.SOUTH);
		this.add(panelHeader, BorderLayout.NORTH);
		// this.add(panelConfig, BorderLayout.NORTH);
		this.add(panelConfig, BorderLayout.CENTER);
		System.out.println("UI Generated");
		tfInputFilePath.setText("<Select File>");

		/*
		 * add(new JLabel("The Accumulated Sum is: ")); tfOutput = new JTextField(10);
		 * tfOutput.setEditable(false); // read-only add(tfOutput);
		 */

		// Allocate an anonymous instance of an anonymous inner class that
		// implements ActionListener as ActionEvent listener
		btnBrowse.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				System.out.println("Step1");
				if (e.getSource() == btnBrowse) {
					System.out.println("Inside Action Listner-Source btnBrows");
					chooser = new JFileChooser(new File(System.getProperty("user.home") + "\\Downloads")); // Downloads
																											// Directory
																											// as
																											// default
					chooser.setDialogTitle("Select Location");
					chooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
					// chooser.setAcceptAllFileFilterUsed(false);

					if (chooser.showOpenDialog(btnBrowse) == JFileChooser.APPROVE_OPTION) {
						fileID = chooser.getSelectedFile().getPath();
						tfInputFilePath.setText(fileID);
						System.out.println(fileID);
					}
				}
				return;
			}
		});
		btnSend.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				System.out.println("Step2");
				String fileNameWithPath = tfInputFilePath.getText();
				System.out.println("fileNameWithPath from tfInputFilePath: " + fileNameWithPath);
				//Read the column ID and send it to split method
				splitFile(fileNameWithPath,Integer.parseInt(tfColumnId.getText()));
				return;
			}
		});
		btnReset.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				tfInputFilePath.setText("<Select File>");
			}
		});

		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE); // Exit program if
														// close-window button
														// clicked
		setTitle("Excel File Splitter"); // "this" Frame sets title
		setSize(350, 200); // "this" Frame sets initial size
		setVisible(true); // "this" Frame shows
	}

	public int splitFile(String fileNameWithPath, int columnId) {
		System.out.println("File to process: " + fileNameWithPath);
		String stringVal="";
		Iterator<Row> itr=null;
		String outputPath="";
		String outputFileName="";
		try {
			
			
			File file = new File(fileNameWithPath); // creating a new file instance
			FileInputStream fis = new FileInputStream(file); // obtaining bytes from the file
			//Create Folder to write output files:
			outputPath = createOutputDir(fileNameWithPath);
			// creating Workbook instance that refers to .xlsx file
			if (fileNameWithPath.endsWith("xlsx")) {
				XSSFWorkbook wb = new XSSFWorkbook(fis);
				XSSFSheet sheet = wb.getSheetAt(0); // creating a Sheet object to retrieve object
				itr = ((Iterable<Row>) sheet).iterator(); // iterating over excel file
			}else if(fileNameWithPath.endsWith("xls")){
				HSSFWorkbook wb=new HSSFWorkbook(fis);   
				HSSFSheet sheet=wb.getSheetAt(0);  
				itr = ((Iterable<Row>) sheet).iterator(); // iterating over excel file
			}else {
				System.out.println("UNSUPPORTED FILE FORMAT");
				return 1;
			}
			while (itr.hasNext()) {
				Row row = itr.next();
				Cell cell = row.getCell(0);
				if(cell.getCellType()==CellType.STRING) 
					stringVal = cell.getStringCellValue(); 
				else if(cell.getCellType()==CellType.NUMERIC) 
					stringVal = String.valueOf(cell.getNumericCellValue());
				System.out.println(stringVal);
				System.out.println("");
				//Create File Name to write this data row
				outputFileName = "SplitFile_" + stringVal + ".xlsx";
				//Call method to write data
				System.out.println("OUTFILE:" + outputPath+ "\\" +outputFileName);
					writeExcelRow(outputPath+ "\\" +outputFileName, row);
				
			}
			fis.close();
		
		} catch (Exception e) {
			e.printStackTrace();
		}

		return 0;
	}
	
	//method to create an output directory in the path where file is present
	public String createOutputDir(String fileNameWithPath) {
		String outputDir="";
		String timeStamp = new SimpleDateFormat("yyyy.MM.dd.HH.mm.ss").format(new java.util.Date());
		String baseDir=fileNameWithPath.substring(0, fileNameWithPath.lastIndexOf("\\"));
		System.out.println(baseDir);
		//Decide a name for the output directory
		outputDir=baseDir + "\\" + "output_"+timeStamp;
		System.out.println(outputDir);
		//Create the output directory
		new File(outputDir).mkdirs();
		return outputDir;
	}
	//method to write a row into a file name
	public boolean writeExcelRow (String fileNameWithPath,Row rowtowrite ) {
		int lastRow=0;
		XSSFRow newRow=null;
		//Row rowtowrite=null;
		XSSFWorkbook workbook = null;
		XSSFSheet sheet = null;

		try {
			//rowtowrite = (XSSFRow)((Row) inprow);
			File outFile=new File(fileNameWithPath);
			
	        //Check whether file already exist
            if(!(outFile.exists())) {
            	outFile.createNewFile();
            	//Create Workbook and Sheet
    			workbook = new XSSFWorkbook();
    	        sheet = workbook.createSheet("Sheet 1");
    	        lastRow=0;
            }
            else {
            		System.out.println("File Exists");
            		//Read the file to see which row shall we write data into
            		FileInputStream fis = new FileInputStream(outFile);
            		workbook = new XSSFWorkbook(fis);
    				sheet = workbook.getSheetAt(0); // creating a Sheet object to retrieve object
    				lastRow=sheet.getPhysicalNumberOfRows();
    				System.out.println("Last Row:" + lastRow);
    				//Close file input stream
    				fis.close();
            }
            //Create New Row
            newRow = sheet.createRow(lastRow+1);
            //Assign the data row from the parent method into the row we just created
            for (int i = 0; i < rowtowrite.getLastCellNum(); i++) {
                // Grab a copy of the old/new cell
                Cell oldCell = rowtowrite.getCell(i);
                XSSFCell newCell = newRow.createCell(i);
                // If the old cell is null jump to next cell
                if (oldCell == null) {
                	System.out.println("OLD CELL VALUE EMPTY FOR CELL: "+ i + ";"+ rowtowrite);
                    continue;
                }

              
                // If there is a cell comment, copy
                if (oldCell.getCellComment() != null) {
                    newCell.setCellComment(oldCell.getCellComment());
                }

                // If there is a cell hyperlink, copy
                if (oldCell.getHyperlink() != null) {
                    newCell.setHyperlink(oldCell.getHyperlink());
                }

                // Set the cell data type
                newCell.setCellType(oldCell.getCellType());

                // Set the cell data value
                switch (oldCell.getCellType()) {
                case BLANK:// Cell.CELL_TYPE_BLANK:
                    newCell.setCellValue(oldCell.getStringCellValue());
                    System.out.println("Cell Value Set:" + oldCell.getStringCellValue());
                    break;
                case BOOLEAN:
                    newCell.setCellValue(oldCell.getBooleanCellValue());
                    System.out.println("Cell Value Set:" + oldCell.getBooleanCellValue());
                    break;
                case FORMULA:
                    newCell.setCellFormula(oldCell.getCellFormula());
                    System.out.println("Cell Value Set:" + oldCell.getCellFormula());
                    break;
                case NUMERIC:
                    newCell.setCellValue(oldCell.getNumericCellValue());
                    System.out.println("Cell Value Set:" + oldCell.getNumericCellValue());
                    break;
                case STRING:
                    newCell.setCellValue(oldCell.getStringCellValue());
                    System.out.println("Cell Value Set:" + oldCell.getStringCellValue());
                    break;
                default:
                    break;
                }
            }

            //Create File output Stream to write
			FileOutputStream fStream = new FileOutputStream(outFile);
	        workbook.write(fStream);
	        //Do cleanup and file close.
	        workbook.close();
	        fStream.close();
	        
		}catch (Exception e) {
			e.printStackTrace();
		}
		return true;
		
	}
	/** The entry main() method */
	public static void main(String[] args) {
		// Run the GUI construction in the Event-Dispatching thread for
		// thread-safety
		SwingUtilities.invokeLater(new Runnable() {
			public void run() {
				
				new splitter();
				// Let the constructor do the job
			}
		});
	}
}