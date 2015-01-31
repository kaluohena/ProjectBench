import java.io.File;
import java.io.IOException;

import javax.swing.JTextArea;

import jxl.Cell;
import jxl.CellType;
import jxl.FormulaCell;
import jxl.Workbook;
import jxl.common.Logger;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.format.CellFormat;
import jxl.read.biff.BiffException;
import jxl.write.Formula;
import jxl.write.Label;
import jxl.write.NumberFormats;
import jxl.write.WritableCell;
import jxl.write.WritableCellFeatures;
import jxl.write.WritableCellFormat;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;


public class ReadWrite {/**
 * The logger
 */
	private static Logger logger = Logger.getLogger(ReadWrite.class);

	/**
	 * The spreadsheet to read in
	 */
	private File inputWorkbook;
	/**
	 * The spreadsheet to output
	 */
	private File outputWorkbook;

	private WritableSheet sheet1;

	private WritableSheet sheet2;

	private JTextArea log;

	private Boolean isCalculateOpen;

	static private final String newline = "\n";

	/**
	 * Constructor
	 * 
	 * @param output 
	 * @param input 
	 */
	public ReadWrite(String input, String output, JTextArea mainLog, Boolean isOpen)
	{
		inputWorkbook = new File(input);
		outputWorkbook = new File(output);
		log = mainLog;
		isCalculateOpen = isOpen;

		logger.setSuppressWarnings(Boolean.getBoolean("jxl.nowarnings"));
		logger.info("Input file:  " + input);    
		logger.info("Output file:  " + output);
	}

	/**
	 * Reads in the inputFile and creates a writable copy of it called outputFile
	 * 
	 * @exception IOException 
	 * @exception BiffException 
	 */
	public void readWrite() throws IOException, BiffException, WriteException
	{
		log.append("Reading..." + newline);
		Workbook w1 = Workbook.getWorkbook(inputWorkbook);

		log.append("Copying..." + newline);
		WritableWorkbook w2 = Workbook.createWorkbook(outputWorkbook, w1);

		modify(w2);

		w2.write();
		w2.close();
		log.append("Done" + newline);
	}

	/*************************************************************************Alice work******************************************************************************************/  
	// Get open point
	protected void getOpenPoint(int thisRow, int anchorRow) throws WriteException, RowsExceededException {
		// Set cell format
		//WritableCellFormat floatFormat = new WritableCellFormat (NumberFormats.FLOAT);
		//floatFormat.setBorder(Border.ALL, BorderLineStyle.THIN);
		String anchorCloseIndex;
		Formula resultFormula = null;

		if (!isCalculateOpen) {
			anchorCloseIndex = "E" + (anchorRow + 1);
		}else {
			anchorCloseIndex = "B" + (anchorRow + 2);
		}

		resultFormula = new Formula(10, thisRow, anchorCloseIndex);
		sheet1.addCell(resultFormula);
	}

	// Get unwinding point
	protected void getUnwindingPoint(int thisRow) throws WriteException, RowsExceededException {
		// Set cell format
		//WritableCellFormat floatFormat = new WritableCellFormat (NumberFormats.FLOAT);
		//floatFormat.setBorder(Border.ALL, BorderLineStyle.THIN);
		String currentCloseIndex;
		Formula resultFormula = null;

		currentCloseIndex = "E" + (thisRow + 1);  

		resultFormula = new Formula(11, thisRow, currentCloseIndex);
		sheet1.addCell(resultFormula);
	}

	// Calculate no loss result (K)
	protected void getNoLossResult(WritableCell thisCell, int anchorRow) throws WriteException, RowsExceededException {
		// Set cell format
		WritableCellFormat floatFormat = new WritableCellFormat (NumberFormats.FLOAT);
		floatFormat.setBorder(Border.ALL, BorderLineStyle.THIN);

		Formula resultFormula = null;
		String openPointIndex = "K" + (thisCell.getRow() + 1);
		String unwindingPointIndex = "L" + (thisCell.getRow() + 1);
		String firstString = thisCell.getContents().substring(0, 1);

		// Don't know the reason but for some cell we get two chars in the cell which have space at end 
		if (firstString.equals("1")) {
			resultFormula = new Formula(12, thisCell.getRow(), unwindingPointIndex+"-"+openPointIndex, floatFormat);
		} else if (firstString.equals("0")) {
			resultFormula = new Formula(12, thisCell.getRow(), openPointIndex+"-"+unwindingPointIndex, floatFormat);
		}
		sheet1.addCell(resultFormula);
	}

	protected void getMaxValue(int thisRow, int conFirstRow) throws WriteException, RowsExceededException {
		// Set cell format
		WritableCellFormat floatFormat = new WritableCellFormat (NumberFormats.FLOAT);
		floatFormat.setBorder(Border.ALL, BorderLineStyle.THIN);

		// construct maximum formula
		String maxFormulaString = "MAX(" + "C" + (conFirstRow + 1) + ":" + "C" + (thisRow + 1);
		Formula maxFormula = new Formula(14, thisRow, maxFormulaString, floatFormat);
		sheet1.addCell(maxFormula);	  
	}

	protected void getMinValue(int thisRow, int conFirstRow) throws WriteException, RowsExceededException {
		// Set cell format
		WritableCellFormat floatFormat = new WritableCellFormat (NumberFormats.FLOAT);
		floatFormat.setBorder(Border.ALL, BorderLineStyle.THIN);

		// construct minimum formula
		String minFormulaString = "MIN(" + "D" + (conFirstRow + 1) + ":" + "D" + (thisRow + 1);
		Formula minFormula = new Formula(15, thisRow, minFormulaString, floatFormat);
		sheet1.addCell(minFormula);
	}

	protected void getMaxProfit(Cell thisCell, int anchorRow) throws WriteException, RowsExceededException {
		// Set cell format
		WritableCellFormat floatFormat = new WritableCellFormat (NumberFormats.FLOAT);
		floatFormat.setBorder(Border.ALL, BorderLineStyle.THIN);

		// construct maximum profit formula
		String maxProfitString = null;
		Formula maxProfitFormula = null;
		String firstString = thisCell.getContents().substring(0, 1);

		if (firstString.equals("1")) {
			maxProfitString = "O" + (thisCell.getRow() + 1) + "-" +"K" + (thisCell.getRow() + 1);
			maxProfitFormula = new Formula(16, thisCell.getRow(), maxProfitString, floatFormat);
		} else if (firstString.equals("0")) {
			maxProfitString =  "K" + (thisCell.getRow() + 1) + "-" + "P" + (thisCell.getRow() + 1);
			maxProfitFormula = new Formula(16, thisCell.getRow(), maxProfitString, floatFormat);
		}

		sheet1.addCell(maxProfitFormula);
	}

	protected void getMaxLoss(Cell thisCell, int anchorRow) throws WriteException, RowsExceededException {
		// Set cell format
		WritableCellFormat floatFormat = new WritableCellFormat (NumberFormats.FLOAT);
		floatFormat.setBorder(Border.ALL, BorderLineStyle.THIN);

		// Construct maximum loss formula
		String maxLossString = null;
		Formula maxLossFormula = null;
		String firstString = thisCell.getContents().substring(0, 1);

		if (firstString.equals("1")) {
			maxLossString = "P" + (thisCell.getRow() + 1) + "-" + "K" + (thisCell.getRow() + 1);
			maxLossFormula = new Formula(17, thisCell.getRow(), maxLossString, floatFormat);
		} else if (firstString.equals("0")) {
			maxLossString = "K" + (thisCell.getRow() + 1) + "-" + "O" + (thisCell.getRow() + 1);
			maxLossFormula = new Formula(17, thisCell.getRow(), maxLossString, floatFormat);
		}

		sheet1.addCell(maxLossFormula);
	}

	protected void getLossResult(int thisRow) throws WriteException, RowsExceededException {
		// Set cell format
		WritableCellFormat floatFormat = new WritableCellFormat (NumberFormats.FLOAT);
		floatFormat.setBorder(Border.ALL, BorderLineStyle.THIN);

		// Construct loss result formula
		String lossResultString = "IF(R" + (thisRow + 1) + "< $B$3,$B$3,M" + (thisRow + 1);
		Formula lossResultFormula = new Formula(13, thisRow, lossResultString, floatFormat);

		sheet1.addCell(lossResultFormula);
	}

	protected void getMaxEarn(Cell thisCell, int conFirstRow) throws WriteException, RowsExceededException {
		// Set cell format
		WritableCellFormat floatFormat = new WritableCellFormat (NumberFormats.FLOAT);
		floatFormat.setBorder(Border.ALL, BorderLineStyle.THIN);

		// Construct loss result formula
		String maxEarnString = null;
		Formula maxEarnFormula = null;
		String firstString = thisCell.getContents().substring(0, 1);

		if (firstString.equals("1")) {
			for (int i = conFirstRow; i < (thisCell.getRow() + 1); i++) {
				maxEarnString = "C" + (i + 1) + "- K" + (thisCell.getRow() + 1);
				maxEarnFormula = new Formula(18, i, maxEarnString, floatFormat);
				sheet1.addCell(maxEarnFormula);
			}
		} else if (firstString.equals("0")) {
			for (int i = conFirstRow; i < (thisCell.getRow() + 1); i++) {
				maxEarnString =  "K" + (thisCell.getRow() + 1) + "- D" + (i + 1);
				maxEarnFormula = new Formula(18, i, maxEarnString, floatFormat);
				sheet1.addCell(maxEarnFormula);
			}
		}

	}

	protected void getMinEarn(Cell thisCell, int conFirstRow) throws WriteException, RowsExceededException {
		// Set cell format
		WritableCellFormat floatFormat = new WritableCellFormat (NumberFormats.FLOAT);
		floatFormat.setBorder(Border.ALL, BorderLineStyle.THIN);

		// Construct loss result formula
		String minEarnString = null;
		Formula minEarnFormula = null;
		String firstString = thisCell.getContents().substring(0, 1);

		if (firstString.equals("1")) {
			for (int i = conFirstRow; i < (thisCell.getRow() + 1); i++) {
				minEarnString = "D" + (i + 1) + "- K" + (thisCell.getRow() + 1);
				minEarnFormula = new Formula(19, i, minEarnString, floatFormat);
				sheet1.addCell(minEarnFormula);
			}
		} else if (firstString.equals("0")) {
			for (int i = conFirstRow; i < (thisCell.getRow() + 1); i++) {
				minEarnString =  "K" + (thisCell.getRow() + 1) + "- C" + (i + 1);
				minEarnFormula = new Formula(19, i, minEarnString, floatFormat);
				sheet1.addCell(minEarnFormula);
			}
		}
	}

	protected void getCloseEarn(Cell thisCell, int conFirstRow) throws WriteException, RowsExceededException {
		// Set cell format
		WritableCellFormat floatFormat = new WritableCellFormat (NumberFormats.FLOAT);
		floatFormat.setBorder(Border.ALL, BorderLineStyle.THIN);

		// Construct loss result formula
		String closeEarnString = null;
		Formula closeEarnFormula = null;
		FormulaCell formulaCell = null;
		String firstString = thisCell.getContents().substring(0, 1);

		if (firstString.equals("1")) {
			for (int i = conFirstRow; i < (thisCell.getRow() + 1); i++) {
				closeEarnString = "E" + (i + 1) + "- K" + (thisCell.getRow() + 1);
				closeEarnFormula = new Formula(20, i, closeEarnString, floatFormat);
				sheet1.addCell(closeEarnFormula);
			}
		} else if (firstString.equals("0")) {
			for (int i = conFirstRow; i < (thisCell.getRow() + 1); i++) {
				closeEarnString =  "K" + (thisCell.getRow() + 1) + "- E" + (i + 1);
				closeEarnFormula = new Formula(20, i, closeEarnString, floatFormat);
				sheet1.addCell(closeEarnFormula);
			}
		}
	}

	//·À²¨µÌ1
	protected void getPreventWave1(int thisRow, int anchorRow ) throws WriteException, RowsExceededException {
		// Set cell format
		WritableCellFormat floatFormat = new WritableCellFormat ();
		floatFormat.setBorder(Border.ALL, BorderLineStyle.THIN);

		String ifMaxFormulaString = "IF(MAX(" + "S" + (anchorRow + 2) + ":" + "S" + (thisRow + 1) + ") >= $C$3, 1, 0";
		Formula ifMaxFormula = new Formula(21, thisRow, ifMaxFormulaString, floatFormat);
		sheet1.addCell(ifMaxFormula);
	}

	//·À²¨µÌ2
	protected void getPreventWave2(int thisRow, int anchorRow ) throws WriteException, RowsExceededException {
		// Set cell format
		WritableCellFormat floatFormat = new WritableCellFormat ();
		floatFormat.setBorder(Border.ALL, BorderLineStyle.THIN);

		String ifMaxFormulaString = "IF(MAX(" + "S" + (anchorRow + 2) + ":" + "S" + (thisRow + 1) + ") >= $D$3, 2, 0";
		Formula ifMaxFormula = new Formula(22, thisRow, ifMaxFormulaString, floatFormat);
		sheet1.addCell(ifMaxFormula);
	}

	//·À²¨µÌ3
	protected void getPreventWave3(int thisRow, int anchorRow) throws WriteException, RowsExceededException {
		// Set cell format
		WritableCellFormat floatFormat = new WritableCellFormat();
		floatFormat.setBorder(Border.ALL, BorderLineStyle.THIN);

		String ifMaxFormulaString = "IF(MAX(" + "S" + (anchorRow + 2) + ":" + "S" + (thisRow + 1) + ") >= $E$3, 3, 0";
		Formula ifMaxFormula = new Formula(23, thisRow, ifMaxFormulaString, floatFormat);
		sheet1.addCell(ifMaxFormula);
	}
	
	//·À²¨µÌ4
	protected void getPreventWave4(int thisRow, int anchorRow) throws WriteException, RowsExceededException {
		// Set cell format
		WritableCellFormat floatFormat = new WritableCellFormat();
		floatFormat.setBorder(Border.ALL, BorderLineStyle.THIN);
		
		String ifMaxFormulaString = "IF(MAX(" + "S" + (anchorRow + 2) + ":" + "S" + (thisRow + 1) + ") >= $F$3, 4, 0";
		Formula ifMaxFormula = new Formula(24, thisRow, ifMaxFormulaString, floatFormat);
		sheet1.addCell(ifMaxFormula);
	}

	//·À²¨µÌ¸ú×Ù1
	protected void getPreventWaveTrack1(int thisRow) throws WriteException, RowsExceededException {
		// Set cell format IF(Vx=0,0,IF(AND(Vx=1,Tx>=0),1,-1))
		WritableCellFormat floatFormat = new WritableCellFormat ();
		floatFormat.setBorder(Border.ALL, BorderLineStyle.THIN);

		String ifMaxFormulaString = "IF(V" + (thisRow + 1) + " = 0, 0, IF(AND(V" + (thisRow + 1) + " = 1, T" + (thisRow + 1) + " >= 0), 1, -1";
		Formula ifMaxFormula = new Formula(25, thisRow, ifMaxFormulaString, floatFormat);
		sheet1.addCell(ifMaxFormula);
	}

	//·À²¨Ìá¸ú×Ù2
	protected void getPreventWaveTrack2(int thisRow, int anchorRow) throws WriteException, RowsExceededException {
		// Set cell format IF(Wx=0,0,IF(AND(Wx=2,Tx>=MAX($S$y:Sx)*$D$4),2,-2))
		WritableCellFormat floatFormat = new WritableCellFormat ();
		floatFormat.setBorder(Border.ALL, BorderLineStyle.THIN);

		String ifMaxFormulaString = "IF(W" + (thisRow + 1) + " = 0, 0, IF(AND(W" + (thisRow + 1) + " = 2, T" + (thisRow + 1) + " >= MAX(S" + (anchorRow + 2) + ": S" + (thisRow + 1) + ") * $D$4), 2, -2))";
		Formula ifMaxFormula = new Formula(26, thisRow, ifMaxFormulaString, floatFormat);
		sheet1.addCell(ifMaxFormula);
	}

	//·À²¨Ìá¸ú×Ù3
	protected void getPreventWaveTrack3(int thisRow, int anchorRow) throws WriteException, RowsExceededException {
		// Set cell format IF(X11=0,0,IF(AND(X11=3,T11>=MAX($S$9:S11)*$E$4),3,-3))
		WritableCellFormat floatFormat = new WritableCellFormat ();
		floatFormat.setBorder(Border.ALL, BorderLineStyle.THIN);

		String ifMaxFormulaString = "IF(X" + (thisRow + 1) + " = 0, 0, IF(AND(X" + (thisRow + 1) + " = 3, T" + (thisRow + 1) + " >= MAX(S" + (anchorRow + 2) + ": S" + (thisRow + 1) + ") * $E$4), 3, -3))";
		Formula ifMaxFormula = new Formula(27, thisRow, ifMaxFormulaString, floatFormat);
		sheet1.addCell(ifMaxFormula);
	}
	
	//·À²¨µÌ¸ú×Ù4
	protected void getPreventWaveTrack4(int thisRow, int anchorRow) throws WriteException, RowsExceededException {
		// Set cell format IF(X11=0,0,IF(AND(X11=3,T11>=MAX($S$9:S11)*$E$4),3,-3))
		WritableCellFormat floatFormat = new WritableCellFormat ();
		floatFormat.setBorder(Border.ALL, BorderLineStyle.THIN);

		String ifMaxFormulaString = "IF(Y" + (thisRow + 1) + " = 0, 0, IF(AND(Y" + (thisRow + 1) + " = 4, T" + (thisRow + 1) + " >= MAX(S" + (anchorRow + 2) + ": S" + (thisRow + 1) + ") * $F$4), 4, -4))";
		Formula ifMaxFormula = new Formula(28, thisRow, ifMaxFormulaString, floatFormat);
		sheet1.addCell(ifMaxFormula);
	}

	protected void addTitle(int titleRow) throws WriteException, RowsExceededException {
		WritableCell thisTitleCell = null;
		CellFormat thisTitleFormat = null;
		Label thisTitle = null;

		// only need to copy column 0 and 9 to the maximum column
		// First add the first title
		thisTitleCell = sheet1.getWritableCell(0, titleRow);
		thisTitleFormat = thisTitleCell.getCellFormat();
		thisTitle = new Label(0, titleRow, thisTitleCell.getContents(), thisTitleFormat);
		sheet2.addCell(thisTitle);
		// Add title from column 9 to the maximum column
		for (int i = 9; i < sheet1.getColumns(); i++) {
			thisTitleCell = sheet1.getWritableCell(i, titleRow);
			thisTitleFormat = thisTitleCell.getCellFormat();
			thisTitle = new Label(i - 8, titleRow, thisTitleCell.getContents(), thisTitleFormat);
			sheet2.addCell(thisTitle);

		}
	}

	protected void addTimeColumn(int thisRow, WritableCell thisCell) throws WriteException, RowsExceededException {
		// due to the type of A column is label, so we add label here
		Label ALabel = new Label(0, thisRow, thisCell.getContents(), thisCell.getCellFormat());
		sheet2.addCell(ALabel);
	}

	protected void AddDataColumn(int thisRow, int thisColumn, WritableCell thisCell) throws WriteException, RowsExceededException {
		Label thisLabel = new Label(thisColumn, thisRow, thisCell.getContents());
		sheet2.addCell(thisLabel);
	}
	/*************************************************************************Alice work******************************************************************************************/
	/**
	 * If the inputFile was the test spreadsheet, then it modifies certain fields
	 * of the writable copy
	 * 
	 * @param w 
	 */
	private void modify(WritableWorkbook w) throws WriteException
	{
		logger.info("Modifying...");

		sheet1 = w.getSheet(0);

		CellFormat cf = null;
		Label l = null;
		WritableCellFeatures wcf = null;
		WritableCell cell = null;

		// Change the format of cell B4 to be emboldened
		/*cell = sheet.getWritableCell(1,3);
	    WritableFont bold = new WritableFont(WritableFont.ARIAL, 
	                                         WritableFont.DEFAULT_POINT_SIZE, 
	                                         WritableFont.BOLD);
	    cf = new WritableCellFormat(bold);
	    cell.setCellFormat(cf);

	    // Change the format of cell B5 to be underlined
	    cell = sheet.getWritableCell(1,4);
	    WritableFont underline = new WritableFont(WritableFont.ARIAL,
	                                              WritableFont.DEFAULT_POINT_SIZE,
	                                              WritableFont.NO_BOLD,
	                                              false,
	                                              UnderlineStyle.SINGLE);
	    cf = new WritableCellFormat(underline);
	    cell.setCellFormat(cf);

	    // Change the point size of cell B6 to be 10 point
	    cell = sheet.getWritableCell(1,5);
	    WritableFont tenpoint = new WritableFont(WritableFont.ARIAL, 10);
	    cf = new WritableCellFormat(tenpoint);
	    cell.setCellFormat(cf);

	    // Change the contents of cell B7 to read "Label - mod"
	    cell = sheet.getWritableCell(1,6);
	    if (cell.getType() == CellType.LABEL)
	    {
	      Label lc = (Label) cell;
	      lc.setString(lc.getString() + " - mod");
	    }

	    // Change cell B10 to display 7 dps
	    cell = sheet.getWritableCell(1,9);
	    NumberFormat sevendps = new NumberFormat("#.0000000");
	    cf = new WritableCellFormat(sevendps);
	    cell.setCellFormat(cf);


	    // Change cell B11 to display in the format 1e4
	    cell = sheet.getWritableCell(1,10);
	    NumberFormat exp4 = new NumberFormat("0.####E0");
	    cf = new WritableCellFormat(exp4);
	    cell.setCellFormat(cf);

	    // Change cell B12 to be normal display
	    cell = sheet.getWritableCell(1,11);
	    cell.setCellFormat(WritableWorkbook.NORMAL_STYLE);

	    // Change the contents of cell B13 to 42
	    cell = sheet.getWritableCell(1,12);
	    if (cell.getType() == CellType.NUMBER)
	    {
	      Number n = (Number) cell;
	      n.setValue(42);
	    }

	    // Add 0.1 to the contents of cell B14
	    cell = sheet.getWritableCell(1,13);
	    if (cell.getType() == CellType.NUMBER)
	    {
	      Number n = (Number) cell;
	      n.setValue(n.getValue() + 0.1);
	    }

	    // Change the date format of cell B17 to be a custom format
	    cell = sheet.getWritableCell(1,16);
	    DateFormat df = new DateFormat("dd MMM yyyy HH:mm:ss");
	    cf = new WritableCellFormat(df);
	    cell.setCellFormat(cf);

	    // Change the date format of cell B18 to be a standard format
	    cell = sheet.getWritableCell(1,17);
	    cf = new WritableCellFormat(DateFormats.FORMAT9);
	    cell.setCellFormat(cf);

	    // Change the date in cell B19 to be 18 Feb 1998, 11:23:28
	    cell = sheet.getWritableCell(1,18);
	    if (cell.getType() == CellType.DATE)
	    {
	      DateTime dt = (DateTime) cell;
	      Calendar cal = Calendar.getInstance();
	      cal.set(1998, 1, 18, 11, 23, 28);
	      Date d = cal.getTime();
	      dt.setDate(d);
	    }

	    // Change the value in B23 to be 6.8.  This should recalculate the 
	    // formula
	    cell = sheet.getWritableCell(1,22);
	    if (cell.getType() == CellType.NUMBER)
	    {
	      Number n = (Number) cell;
	      n.setValue(6.8);
	    }

	    // Change the label in B30.  This will have the effect of making
	    // the original string unreferenced
	    cell = sheet.getWritableCell(1, 29);
	    if (cell.getType() == CellType.LABEL)
	    {
	      l = (Label) cell;
	      l.setString("Modified string contents");
	    }
	    // Insert a new row (number 35)
	    sheet.insertRow(34);

	    // Delete row 38 (39 after row has been inserted)
	    sheet.removeRow(38);

	    // Insert a new column (J)
	    sheet.insertColumn(9);

	    // Remove a column (L - M after column has been inserted)
	    sheet.removeColumn(11);

	    // Remove row 44 (contains a hyperlink), and then insert an empty
	    // row just to keep the numbers consistent
	    sheet.removeRow(43);
	    sheet.insertRow(43);

	    // Modify the hyperlinks
	    WritableHyperlink hyperlinks[] = sheet.getWritableHyperlinks();

	    for (int i = 0; i < hyperlinks.length; i++)
	    {
	      WritableHyperlink wh = hyperlinks[i];
	      if (wh.getColumn() == 1 && wh.getRow() == 39)
	      {
	        try
	        {
	          // Change the hyperlink that begins in cell B40 to be a different API
	          wh.setURL(new URL("http://www.andykhan.com/jexcelapi/index.html"));
	        }
	        catch (MalformedURLException e)
	        {
	          logger.warn(e.toString());
	        }
	      }
	      else if (wh.getColumn() == 1 && wh.getRow() == 40)
	      {
	        wh.setFile(new File("../jexcelapi/docs/overview-summary.html"));
	      }
	      else if (wh.getColumn() == 1 && wh.getRow() == 41)
	      {
	        wh.setFile(new File("d:/home/jexcelapi/docs/jxl/package-summary.html"));
	      }
	      else if (wh.getColumn() == 1 && wh.getRow() == 44)
	      {
	        // Remove the hyperlink at B45
	        sheet.removeHyperlink(wh);
	      }
	    }

	    // Change the background of cell F31 from blue to red
	    WritableCell c = sheet.getWritableCell(5,30);
	    WritableCellFormat newFormat = new WritableCellFormat(c.getCellFormat());
	    newFormat.setBackground(Colour.RED);
	    c.setCellFormat(newFormat);

	    // Modify the contents of the merged cell
	    l = new Label(0, 49, "Modified merged cells");
	    sheet.addCell(l);

	    // Modify the chart data
	    Number n = (Number) sheet.getWritableCell(0, 70);
	    n.setValue(9);

	    n = (Number) sheet.getWritableCell(0, 71);
	    n.setValue(10);

	    n = (Number) sheet.getWritableCell(0, 73);
	    n.setValue(4);*/

		/*************************************************************************Alice work******************************************************************************************/   
		WritableCell curFormulaCell = null;
		WritableCell nextFormulaCell = null;
		WritableCellFormat cellFormat = null;
		CellFormat format = null;
		Cell tempCell = null;
		WritableCell testCell = null;
		int startRow = 0;
		int anchorRow = 0;

		// Get the first number in J column
		for (int i = 0; i < sheet1.getRows() - 1; i++) {
			tempCell = sheet1.getCell(9, i);
			if (tempCell.getType() == CellType.NUMBER_FORMULA || tempCell.getType() == CellType.NUMBER) {
				startRow = i;
				break;
			}
		}
		logger.info("the real start row: " + startRow + 1);
		logger.info("total row: "+ sheet1.getRows());
		if (!isCalculateOpen) {
			anchorRow = startRow; 
		} else {
			anchorRow = startRow - 1;
		}

		curFormulaCell = sheet1.getWritableCell(9, startRow);

		// Here we got two index system: Cell index and real excel index
		// Apart from the formula, we all use cell index
		for (int i = (startRow + 1); i < (sheet1.getRows() + 1); i++) {
			nextFormulaCell = sheet1.getWritableCell(9, i);
			// set color for current cell
			/* firstString = curFormulaCell.getContents().substring(0, 1);
	    	if (firstString.equals("1")) {
	    		cellFormat = new WritableCellFormat(curFormulaCell.getCellFormat());
	    		cellFormat.setBackground(Colour.LIGHT_GREEN);
	    		curFormulaCell.setCellFormat(cellFormat);
	    	}*/

			String curFormulaCellString = curFormulaCell.getContents();
			String nextFormulaCellString = nextFormulaCell.getContents(); 

			getPreventWave1(curFormulaCell.getRow(), anchorRow);
			getPreventWave2(curFormulaCell.getRow(), anchorRow);
			getPreventWave3(curFormulaCell.getRow(), anchorRow);
			getPreventWave4(curFormulaCell.getRow(), anchorRow);
			getPreventWaveTrack1(curFormulaCell.getRow());
			getPreventWaveTrack2(curFormulaCell.getRow(), anchorRow);
			getPreventWaveTrack3(curFormulaCell.getRow(), anchorRow);
			getPreventWaveTrack4(curFormulaCell.getRow(), anchorRow);

			if (!curFormulaCellString.replaceAll("\\s", "").equals(nextFormulaCellString.replaceAll("\\s", ""))) {
				log.append("Row "+i+" : current cell: "+curFormulaCell.getContents()+", next cell contents: "+nextFormulaCell.getContents() + newline);
				format = curFormulaCell.getCellFormat();
				// Get open point
				getOpenPoint(curFormulaCell.getRow(), anchorRow);

				// Get unwinding point
				getUnwindingPoint(curFormulaCell.getRow());

				// Calculate result without loss
				getNoLossResult(curFormulaCell, anchorRow);

				// Calculate the maximum value
				// the first excel row for the succession equals anchorRow + 1
				getMaxValue(curFormulaCell.getRow(), anchorRow + 1);

				// Calculate the minimum value
				getMinValue(curFormulaCell.getRow(), anchorRow + 1);

				// Calculate the maximum profit
				getMaxProfit(curFormulaCell, anchorRow);

				// Calculate the maximum loss
				getMaxLoss(curFormulaCell, anchorRow);

				// Calculate the result with loss
				getLossResult(curFormulaCell.getRow());

				// Calculate the max earn
				getMaxEarn(curFormulaCell, anchorRow + 1);

				// Calculate the min earn
				getMinEarn(curFormulaCell, anchorRow + 1);

				// Calculate the close earn
				getCloseEarn(curFormulaCell, anchorRow + 1);

				anchorRow = i - 1;
			}

			curFormulaCell = nextFormulaCell;
		}

		// create another sheet1
		/*w.createSheet("Result", 1);
	    sheet2 = w.getSheet(1);
	    WritableCell tempTitleCell = null;

	    // get the table title 
	    int lastColumn  = sheet1.getColumns() - 1;
	    int titleRow = 0;
	    for (int i = 0; i < sheet1.getRows(); i++) {
	    	  tempTitleCell = sheet1.getWritableCell(lastColumn, i);
	    	  if (tempTitleCell.getType() == CellType.LABEL) {
	    		  titleRow = i;
	    		  break;
	    	  }
	    }

	    // Write the table title in shee1 to sheet2
	    addTitle(titleRow);
	    // Find the row that has value in unwinding column
	    Label timeLabel = null;
	    Number thisNumber = null;
	    WritableCell thisTimeCell = null;
	    WritableCell thisDataCell = null;
	    int currentRow2 = titleRow + 1;

	    for (int i = startRow; i < sheet1.getRows(); i++) {
	    	thisDataCell = sheet1.getWritableCell(10, i);
	    	// Using unwinding column to get processed data row 
	    	if (!thisDataCell.getContents().isEmpty()) {
	    		// Copy the column A in processed data row to sheet2
	    		thisTimeCell = sheet1.getWritableCell(0, i);
	    		addTimeColumn(currentRow2, thisTimeCell);    		

	    		for (int j = 9; j < sheet1.getColumns(); j++) {
	    			thisDataCell = sheet1.getWritableCell(j, i);
	    			if (j == 10) {
	    				Formula thisFormula = (Formula)thisDataCell;
	    				logger.info("data: "+thisFormula.getContents());
	    			}
	    			AddDataColumn(currentRow2, j - 8, thisDataCell);
	    		}
	    		currentRow2++;
	    	}
	    }*/


		/*************************************************************************Alice work******************************************************************************************/
		// Add in a cross sheet formula
		/*Formula f = new Formula(1, 80, "ROUND(COS(original!B10),2)");
	    sheet.addCell(f);

	    // Add in a formula from the named cells
	    f = new Formula(1, 83, "value1+value2");
	    sheet.addCell(f);

	    // Add in a function formula using named cells
	    f = new Formula(1, 84, "AVERAGE(value1,value1*4,value2)");
	    sheet.addCell(f);

	    // Copy sheet 1 to sheet 3
	    //     w.copySheet(0, "copy", 2);

	    // Use the cell deep copy feature
	    Label label = new Label(0, 88, "Some copied cells", cf);
	    sheet.addCell(label);

	    label = new Label(0,89, "Number from B9");
	    sheet.addCell(label);

	    WritableCell wc = sheet.getWritableCell(1, 9).copyTo(1,89);
	    sheet.addCell(wc);

	    label = new Label(0, 90, "Label from B4 (modified format)");
	    sheet.addCell(label);

	    wc = sheet.getWritableCell(1, 3).copyTo(1,90);
	    sheet.addCell(wc);

	    label = new Label(0, 91, "Date from B17");
	    sheet.addCell(label);

	    wc = sheet.getWritableCell(1, 16).copyTo(1,91);
	    sheet.addCell(wc);

	    label = new Label(0, 92, "Boolean from E16");
	    sheet.addCell(label);

	    wc = sheet.getWritableCell(4, 15).copyTo(1,92);
	    sheet.addCell(wc);

	    label = new Label(0, 93, "URL from B40");
	    sheet.addCell(label);

	    wc = sheet.getWritableCell(1, 39).copyTo(1,93);
	    sheet.addCell(wc);

	    // Add some numbers for the formula copy
	    for (int i = 0 ; i < 6; i++)
	    {
	      Number number = new Number(1,94+i, i + 1 + i/8.0);
	      sheet.addCell(number);
	    }

	    label = new Label(0,100, "Formula from B27");
	    sheet.addCell(label);

	    wc = sheet.getWritableCell(1, 26).copyTo(1,100);
	    sheet.addCell(wc);

	    label = new Label(0,101, "A brand new formula");
	    sheet.addCell(label);

	    Formula formula = new Formula(1, 101, "SUM(B94:B96)");
	    sheet.addCell(formula);

	    label = new Label(0,102, "A copy of it");
	    sheet.addCell(label);

	    wc = sheet.getWritableCell(1,101).copyTo(1, 102);
	    sheet.addCell(wc);

	    // Remove the second image from the sheet
	    WritableImage wi = sheet.getImage(1);
	    sheet.removeImage(wi);

	    wi = new WritableImage(1, 116, 2, 9, 
	                           new File("resources/littlemoretonhall.png"));
	    sheet.addImage(wi); */

		// Add a list data validations
		/*label = new Label(0, 151, "Added drop down validation");
	    sheet.addCell(label);

	    Blank b = new Blank(1, 151);
	    wcf = new WritableCellFeatures();
	    ArrayList al = new ArrayList();
	    al.add("The Fellowship of the Ring");
	    al.add("The Two Towers");
	    al.add("The Return of the King");
	    wcf.setDataValidationList(al);
	    b.setCellFeatures(wcf);
	    sheet.addCell(b);

	    // Add a number data validation
	    label = new Label(0, 152, "Added number validation 2.718 < x < 3.142");
	    sheet.addCell(label);
	    b = new Blank(1,152);
	    wcf = new WritableCellFeatures();
	    wcf.setNumberValidation(2.718, 3.142, wcf.BETWEEN);
	    b.setCellFeatures(wcf);
	    sheet.addCell(b);

	    // Modify the text in the first cell with a comment
	    cell = sheet.getWritableCell(0, 156);
	    l = (Label) cell;
	    l.setString("Label text modified");

	    cell = sheet.getWritableCell(0, 157);
	    wcf = cell.getWritableCellFeatures();
	    wcf.setComment("modified comment text");

	    cell = sheet.getWritableCell(0, 158);
	    wcf = cell.getWritableCellFeatures();
	    wcf.removeComment();

	    // Modify the validation contents of the row 173
	    cell = sheet.getWritableCell(0,172);
	    wcf = cell.getWritableCellFeatures();
	    Range r = wcf.getSharedDataValidationRange();
	    Cell botright = r.getBottomRight();
	    sheet.removeSharedDataValidation(cell);
	    al = new ArrayList();
	    al.add("Stanley Featherstonehaugh Ukridge");
	    al.add("Major Plank");
	    al.add("Earl of Ickenham");
	    al.add("Sir Gregory Parsloe-Parsloe");
	    al.add("Honoria Glossop");
	    al.add("Stiffy Byng");
	    al.add("Bingo Little");
	    wcf.setDataValidationList(al);
	    cell.setCellFeatures(wcf);
	    sheet.applySharedDataValidation(cell, 
	                                    botright.getColumn() - cell.getColumn(),
	                                    1);//botright.getRow() - cell.getRow());*/
	}
}
