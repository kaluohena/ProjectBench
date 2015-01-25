import java.io.File;
import java.io.IOException;
import java.util.ArrayList;

import javax.swing.JTextArea;

import jxl.Cell;
import jxl.Workbook;
import jxl.common.Logger;
import jxl.format.CellFormat;
import jxl.read.biff.BiffException;
import jxl.write.Number;
import jxl.write.NumberFormats;
import jxl.write.WritableCellFormat;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;


public class CopyPaste {
	private static Logger logger = Logger.getLogger(CopyPaste.class);
	private static final String newline = "\n";
	
	private File inputWorkbook;
	private File outputWorkbook;
	private JTextArea log;
	
	public CopyPaste(String input, String output, JTextArea mainLog) {
		inputWorkbook = new File(input);
		outputWorkbook = new File(output);
		log = mainLog;
		
		logger.setSuppressWarnings(Boolean.getBoolean("jxl.nowarnings"));
	    logger.info("Input file:  " + input);    
	    logger.info("Output file:  " + output);
	}
	
	public void copyPaste() throws BiffException, IOException, WriteException {
		logger.info("Reading...");
	    Workbook w1 = Workbook.getWorkbook(inputWorkbook);

	    logger.info("Copying...");
	    WritableWorkbook w2 = Workbook.createWorkbook(outputWorkbook, w1);
	    
	    if (!execute(w2)) {
	    	log.append("Erro" + newline);
	    } else {
	    	w2.write();
		    w2.close();
		    log.append("Done" + newline);
	    } 
	}
	
	protected Boolean execute(WritableWorkbook w) throws RowsExceededException, WriteException {
		logger.info("Copying...");
		
		WritableSheet copySheet = w.getSheet(1);
		WritableSheet pasteSheet = w.getSheet(0);
		Cell copyCalendarCell;
		Cell pasteCalendarCell;
		Cell copyOpenCell;
		Cell copyMaxCell;
		Cell copyMinCell;
		Cell copyCloseCell;
		Cell maxCell = null;
		Cell minCell = null;
		CellFormat cellFormat;
		MyCalendar copyMyCalendar = new MyCalendar();
		WritableCellFormat wCellFormat = new WritableCellFormat(NumberFormats.FLOAT);
		ArrayList<MyCalendar> pasteCalendarArrayList = new ArrayList<MyCalendar>();
		int copySheetStartIndex;
		int pasteSheetStartIndex;
		int startIndex = -1;
		int endIndex  = -1;
		
		//check whether the copySheet has time column
		if (sheetValidate(copySheet) < 0 || sheetValidate(pasteSheet) < 0) {
			//if there is no Calendar column return false
			return false;
		}
				
		copySheetStartIndex = sheetValidate(copySheet) ;
		pasteSheetStartIndex = sheetValidate(pasteSheet);
		//retrieve calendar information from pasteSheet
		for (int i = pasteSheetStartIndex; i < pasteSheet.getRows(); i++) {
			pasteCalendarCell = pasteSheet.getCell(0, i);
			//covert cell content string to calendar class
			if (pasteCalendarCell.getContents() != "") {
				pasteCalendarArrayList.add(StringCalendarConvertion(pasteCalendarCell));
			}		
		}
		
		for (int j = 0; j < pasteCalendarArrayList.size(); j++) {
			for (int i = copySheetStartIndex; i < copySheet.getRows(); i++) {
				copyCalendarCell = copySheet.getCell(0, i);
				//covert cell content string to calendar class
				if (copyCalendarCell.getContents() != "") {
					copyMyCalendar = StringCalendarConvertion(copyCalendarCell);
				} else {	
					break;
				}
				
				//for the special time 13:00 in pasteSheet which correspond to 11:30 in copySheet
				//convert 11:30 to 13:00 in copySheet
				if (copyMyCalendar.getHour() == 11 && copyMyCalendar.getMinute() == 30) {
					copyMyCalendar.setHour(13);
					copyMyCalendar.setMinute(0);
				}
				
				//start to paste content from pasteSheet to copySheet
				if (j - 1 < 0) {
					break;
				} else {
					//get the start index for the period in copySheet
					if (copyMyCalendar.getMonth() == (pasteCalendarArrayList.get(j-1)).getMonth() &&
						copyMyCalendar.getDay() == (pasteCalendarArrayList.get(j-1)).getDay() &&
						copyMyCalendar.getHour() == (pasteCalendarArrayList.get(j-1)).getHour() &&
						copyMyCalendar.getMinute() == (pasteCalendarArrayList.get(j-1)).getMinute()) {						
							if (copyMyCalendar.getYear() == 0 || (copyMyCalendar.getYear() > 0 && copyMyCalendar.getYear() == (pasteCalendarArrayList.get(j-1)).getYear())) {
								startIndex = i + 1;
								maxCell = copySheet.getCell(2, startIndex);
								minCell = copySheet.getCell(3, startIndex);
							} 
					}
					
					//get the end index for the period in copySheet
					if (startIndex > 0) {
						if (copyMyCalendar.getMonth() == (pasteCalendarArrayList.get(j)).getMonth() &&
							copyMyCalendar.getDay() == (pasteCalendarArrayList.get(j)).getDay() &&
							copyMyCalendar.getHour() == (pasteCalendarArrayList.get(j)).getHour() &&
							copyMyCalendar.getMinute() == (pasteCalendarArrayList.get(j)).getMinute()) {
							if (copyMyCalendar.getYear() == 0 || (copyMyCalendar.getYear() > 0 && copyMyCalendar.getYear() == (pasteCalendarArrayList.get(j)).getYear()))
								endIndex = i;
						}
						
						//get max cell and min cell in period
						//because the period start from the next row(i+1) of corresponding time's row in sheet2, so we don't need the content in row i 
						if (i + 1 != startIndex) {
							Cell thisMaxCell = copySheet.getCell(2, i);
							Cell thisMinCell = copySheet.getCell(3, i);
							if (Double.parseDouble(thisMaxCell.getContents()) > Double.parseDouble(maxCell.getContents())) {
								maxCell = thisMaxCell;
							}						
							if (Double.parseDouble(thisMinCell.getContents()) < Double.parseDouble(minCell.getContents())) {
								minCell = thisMinCell;
							}
						}
						
						//copy
						if (startIndex > 0 && endIndex > 0) {
							//copy open cell data to pasteSheet
							copyOpenCell = copySheet.getCell(1, startIndex);
							cellFormat = copyOpenCell.getCellFormat();
							Number openNumber = new Number(1, (pasteSheetStartIndex + j), Double.parseDouble(copyOpenCell.getContents()), cellFormat);
							pasteSheet.addCell(openNumber);
							
							//copy Max cell data to pasteSheet
							copyMaxCell = maxCell;
							cellFormat = copyMaxCell.getCellFormat();
							Number maxNumber = new Number(2, (pasteSheetStartIndex + j), Double.parseDouble(copyMaxCell.getContents()), cellFormat);
							pasteSheet.addCell(maxNumber);
							
							//copy Min column data to pasteSheet
							copyMinCell = minCell;
							cellFormat = copyMinCell.getCellFormat();
							Number minNumber = new Number(3, (pasteSheetStartIndex + j), Double.parseDouble(copyMinCell.getContents()), cellFormat);
							pasteSheet.addCell(minNumber);
							
							//copy Close column data to pasteSheet
							copyCloseCell = copySheet.getCell(4, endIndex);
							cellFormat = copyCloseCell.getCellFormat();
							Number closeNumber = new Number(4, (pasteSheetStartIndex + j), Double.parseDouble(copyCloseCell.getContents()), cellFormat);
							pasteSheet.addCell(closeNumber);
							
							startIndex = -1;
							endIndex = -1;
						}
					}
				}
			}
			if (startIndex != -1 || endIndex != -1) {
				if (startIndex < 0 && endIndex < 0) {
					log.append("cannot find the time in " + (pasteSheetStartIndex + j) + " of sheet1 in sheet2" + newline);
				}
				if (startIndex > 0 && endIndex < 0) {
					log.append("cannot find the time in " + (pasteSheetStartIndex + j + 1) + " of sheet1 in sheet2" + newline);
				}
				break;
			}
		}
		
		
		//after finish executing return true
		return true;
	}
	
	protected int sheetValidate(WritableSheet sheet) {
		Cell calendarCell;
		for (int i = 0; i < sheet.getRows(); i++) {
			calendarCell = sheet.getCell(0, i);
			if (calendarCell.getContents().indexOf("Ê±¼ä") >= 0) {
				return i + 1;
			}
		}
		log.append("Cannot find calendar column" + newline);
		return -1;
	}
	
	protected MyCalendar StringCalendarConvertion(Cell cell) {
		String calendarPattern = ".*/.*-.*";
		String year="", month="";
		if (cell.getContents().matches(calendarPattern)) {
			String[] slashTokens = cell.getContents().split("/");
			for (int i = 0; i < slashTokens.length; i++) {
				String hyphenPattern = ".*-.*";
				if (slashTokens[i].matches(hyphenPattern)) {
					String[] hyphenTokens = slashTokens[i].split("-");
					String[] colonTokens = hyphenTokens[1].split(":");
					//Remove the space from data and time string
					//year
					if (slashTokens.length == 2) { // judge whether the year information added
						month = slashTokens[0].replaceAll("\\s+", "");
					} else if (slashTokens.length == 3) {
						year = slashTokens[0].replaceAll("\\s+", "");
						month = slashTokens[1].replaceAll("\\s+", "");
					}
					//day
					hyphenTokens[0] = hyphenTokens[0].replaceAll("\\s", "");
					//hour
					colonTokens[0] = colonTokens[0].replaceAll("\\s", "");
					//minute
					colonTokens[1] = colonTokens[1].replaceAll("\\s", "");
					
					//create MyCalendar class instance
					MyCalendar myCalendar = new MyCalendar();
					//assign value to myCalendar
					if (!year.isEmpty())
						myCalendar.setYear(Integer.parseInt(year));
					if (!month.isEmpty())
						myCalendar.setMonth(Integer.parseInt(month));
					myCalendar.setDay(Integer.parseInt(hyphenTokens[0]));
					myCalendar.setHour(Integer.parseInt(colonTokens[0]));
					myCalendar.setMinute(Integer.parseInt(colonTokens[1]));
					
					return myCalendar;
				}
			}
		} 
			
		log.append("row " + (cell.getRow() + 1) + " doesn't satisfy calendar format." + newline);
		return null;
	}

}
