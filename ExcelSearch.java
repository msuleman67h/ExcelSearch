/************************
 * Author: msulema5
 * Program: Excel Search for Polypropylen Resin & % GF Information.xlsx
 * Gives the of searched value/ query
 ************************/

import java.awt.EventQueue;
import java.awt.Font;
import java.awt.event.ComponentAdapter;
import java.awt.event.ComponentEvent;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.JTextField;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.table.DefaultTableModel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelSearch {

	// Instance Variables 
	private JFrame frmSearch;
	private JTable searchResult;
	private String fileName = "";
	private String filePath = "";
	private JTextField searchTextField;
	private String[] excelRow = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M"};
	private String keyword = "";

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					ExcelSearch window = new ExcelSearch();
					window.frmSearch.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the application.
	 * @throws Exception 
	 */
	public ExcelSearch() throws Exception {
		initialize();
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		// Frame Properties 
		frmSearch = new JFrame();
		frmSearch.setTitle("Excel Search");
		frmSearch.setBounds(100, 100, 820, 600);
		frmSearch.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frmSearch.getContentPane().setLayout(null);

		/**
		 * JLabel lblFileName
		 * Displays selected file name
		 */
		JLabel lblFileName = new JLabel("No file selected");
		lblFileName.setBounds(430, 10, 247, 23);
		frmSearch.getContentPane().add(lblFileName);

		/**
		 * JButton btnBrowse
		 * Opens file chooser to select file
		 */
		JButton btnBrowse = new JButton("Browse");
		btnBrowse.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseClicked(MouseEvent arg0) {

				/**
				 * JFileChooser chooser
				 * Allows user to select file
				 */
				JFileChooser chooser = new JFileChooser();

				// Applying filter so that user can only select excel files
				FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel Files", "xlsx", "excel");
				chooser.setFileFilter(filter);
				chooser.showOpenDialog(frmSearch);

				// Saves the name and path to file
				File file = chooser.getSelectedFile();
				fileName = file.getName();
				filePath = file.getAbsolutePath();
				lblFileName.setText(fileName);
			}
		});
		btnBrowse.setBounds(685, 10, 89, 23);
		frmSearch.getContentPane().add(btnBrowse);

		JScrollPane scrollPane = new JScrollPane();
		scrollPane.setBounds(10, 44, frmSearch.getWidth()-35, frmSearch.getHeight()-90);

		/**
		 * Component Listener
		 * Checks for changes in frame size and resizes the components inside accordingly 
		 */
		frmSearch.addComponentListener(new ComponentAdapter() {

			public void componentResized(final ComponentEvent e) {
				scrollPane.setBounds(10, 44, frmSearch.getWidth()-35, frmSearch.getHeight()-90);
				lblFileName.setBounds(frmSearch.getWidth()-390, 10, 247, 23);
				btnBrowse.setBounds(frmSearch.getWidth()-115, 10, 89, 23);
			}
		});
		frmSearch.getContentPane().add(scrollPane);

		/**
		 * JTable searchResult
		 * Displays the search result in table similar to excel
		 */
		searchResult = new JTable();
		searchResult.setFont(new Font("Tahoma", Font.PLAIN, 13));
		searchResult.setModel(new DefaultTableModel(
				new Object[][] {
				},
				new String[] {
						"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L"
				}
				));
		scrollPane.setViewportView(searchResult);
		DefaultTableModel tableModel = (DefaultTableModel) searchResult.getModel();

		/**
		 * JTextField searchTextField
		 * Allows user to enter search query
		 */
		searchTextField = new JTextField();
		searchTextField.setToolTipText("Enter search query to search in excel");
		searchTextField.setBounds(10, 10, 315, 23);
		frmSearch.getContentPane().add(searchTextField);
		searchTextField.setColumns(10);

		/**
		 * JButton btnSearch
		 * Searches for query in excel file
		 */
		JButton btnSearch = new JButton("Search");
		btnSearch.addMouseListener(new MouseAdapter() {
			private XSSFWorkbook workbook;

			@Override
			public void mouseClicked(MouseEvent s) {
				// Saves the search query in a string 
				keyword = searchTextField.getText();
				// If either query or file is missing 
				if (lblFileName.getText().equals("No file selected") || keyword.length() == 0) { 
					JOptionPane.showMessageDialog(frmSearch, "No file selected or search query is missing!", 
							"Error", JOptionPane.ERROR_MESSAGE);
				} else {
					// Clears the table if it is not empty 
					if(tableModel.getRowCount() != 0) {
						tableModel.setRowCount(0);
					}
					try {

						FileInputStream excelFile = new FileInputStream(new File(filePath));
						// Creates a workbook that refers to .xls file
						workbook = new XSSFWorkbook(excelFile);
						// Creates a sheet to retrieve the sheet
						XSSFSheet sheet = workbook.getSheetAt(0);
						// Evaluates cell data type
						FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
						// rowHeader, true if the row contains headers 
						boolean rowHeader = false;
						// True if row contains search query 
						boolean rowData = false;
						int arrayIndex = 0;
						for (Row row : sheet) {
							for (Cell cell : row) {
								// Checks for a specific heading
								for(int i = 10; i <= 60; i = i+5) {
									if(evaluator.evaluateInCell(cell).getCellTypeEnum() == CellType.STRING) {
										if(cell.getStringCellValue().equals(" PP  "+ i + "% Glass Fiber Filled,"
												+ "   vendors Comparison  ")) {
											rowData = true;
										}
									}
								}
								// Evaluates the data in cell and decides if it's header or data
								if (evaluator.evaluateInCell(cell).getCellTypeEnum() == CellType.STRING) {
									if (cell.getStringCellValue().equals("Material Trade Name / Grade")
											|| cell.getStringCellValue().equals("Polypropylene with different"
													+ " Glass Fiber %")) {
										// If it's header
										rowHeader = true;
									} else if (cell.getStringCellValue().equalsIgnoreCase(keyword)) {
										// If it's data
										rowData = true;
									}
								}
								// Adds header/data row to an array (excelRow)
								if ((rowHeader == true) && evaluator.evaluateInCell(cell).getCellTypeEnum() 
										== CellType.STRING) {
									excelRow[arrayIndex] = cell.getStringCellValue();
									arrayIndex++;
								} else if (rowData == true) {
									if ((evaluator.evaluateInCell(cell).getCellTypeEnum() == CellType.NUMERIC)){
										excelRow[arrayIndex] = cell.getNumericCellValue() + "";
										arrayIndex++;
									} else if (evaluator.evaluateInCell(cell).getCellTypeEnum() 
											== CellType.STRING) {
										excelRow[arrayIndex] = cell.getStringCellValue();
										arrayIndex++;
									} else {
										excelRow[arrayIndex] = cell.getStringCellValue();
										arrayIndex++;
									}
								}
							}
							// Adds excelRow to table searchResult
							if ((rowHeader == true) || (rowData == true)) {
								tableModel.addRow(new Object[] { excelRow[0], excelRow[1], excelRow[2], 
										excelRow[3],excelRow[4], excelRow[5], excelRow[6], excelRow[7],
										excelRow[8], excelRow[9],excelRow[10], excelRow[11], excelRow[12] });
								for (int i = 0; i < excelRow.length; i++) {
									excelRow[i] = "";
								}
							}
								// Resets all values to evaluate next row
								arrayIndex = 0;
								rowHeader = false;
								rowData = false;
							}
						} catch (FileNotFoundException e) {
							e.printStackTrace();
						} catch (IOException i) {
							i.printStackTrace();
						}
					}
				}
			});
		btnSearch.setBounds(335, 10, 89, 23);
		frmSearch.getContentPane().add(btnSearch);
		}
	}
