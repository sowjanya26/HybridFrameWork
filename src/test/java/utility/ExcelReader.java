package utility;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Calendar;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;




public class ExcelReader {

	/**
	 * @param args
	 * @throws IOException 
	 * @throws InvalidFormatException 
	 * 
	 */
	public static String path;
	public static File file =null;
	public static FileInputStream fi =null;
	public static Workbook book =null;
	public static Sheet sheet =null;
	public static Row row =null; 
	public static Cell cell =null;
	
	public ExcelReader(String path) throws InvalidFormatException, IOException{
		this.path = path;
		try{
			file = new File(path);
			fi = new FileInputStream(file);
			book = WorkbookFactory.create(fi);
			Sheet sheet0 = book.getSheetAt(0);
			fi.close();
			
		}catch(Exception e){
			e.printStackTrace();
		}
		
	}
	
	//returns number of rows
	public int getRowCount(String sheetName){
		
		int index = book.getSheetIndex(sheetName);
		if(index ==-1){
			return 0;
		}else{
			sheet=book.getSheetAt(index);
			return sheet.getLastRowNum()+1;
		}
		
	}
	
	//returns number of Columns
		public int getColumnCount(String sheetName){
			if(!isSheetExists(sheetName))
				return -1;
			sheet = book.getSheet(sheetName);
			row = sheet.getRow(0);
			if(row==null)
				return -1;
			else
				return row.getLastCellNum();
			
		}
		
		//find whether sheet exists or not
		public boolean isSheetExists(String sheetName){
			int index = book.getSheetIndex(sheetName);
			if(index==-1){
				index = book.getSheetIndex(sheetName.toUpperCase());
				if(index==-1)
					return false;
				else 
					return true;
			}else
				return true;
			
		}
	
	//returns the data from the cell using Column Name & row Number
	public String getCellData(String sheetName, String colName,int rowNum){
		try{
			if(rowNum <=0)
				return "";
			int index = book.getSheetIndex(sheetName);
			int colNum = -1;
			if (index ==-1)
				return "";
			sheet = book.getSheetAt(index);
			row = sheet.getRow(0);
			for(int i=0; i<row.getLastCellNum();i++){
				if(row.getCell(i).getStringCellValue().trim().equals(colName))
					colNum = i;
			}
			if(colNum ==-1)
				return "";
			sheet = book.getSheet(sheetName);
			row = sheet.getRow(rowNum-1);
			if (row==null)
				return "";
			cell = row.getCell(colNum);			
			if(cell==null)
				return "";
			
			if(cell.getCellType()==Cell.CELL_TYPE_STRING)
				return cell.getStringCellValue();
			else if(cell.getCellType()==Cell.CELL_TYPE_NUMERIC || cell.getCellType()==Cell.CELL_TYPE_FORMULA){
				String cellText = String.valueOf(cell.getNumericCellValue());
				if (HSSFDateUtil.isCellDateFormatted(cell)){
					double d = cell.getNumericCellValue();
					Calendar cal = Calendar.getInstance();
					cal.setTime(HSSFDateUtil.getJavaDate(d));
					cellText = (String.valueOf(cal.get(Calendar.YEAR))).substring(2);
					cellText = cal.get(Calendar.MONTH)+1+"/"+
							cal.get(Calendar.DAY_OF_MONTH)+"/"+
							cellText;
				}
				return cellText;
			}else if(cell.getCellType()==Cell.CELL_TYPE_BLANK)
				return "";
			else
				return String.valueOf(cell.getBooleanCellValue());
				
			
			
		} catch(Exception e){
			e.printStackTrace();
			return "Row Number "+rowNum + "or Column Number "+colName +"doesn't exist in Excel.";
			
		}
		
	}

	//returns the data from the cell using Column Number & row Number
	public String getCellData(String sheetName, int colNum,int rowNum){
		try{
			if(rowNum <=0)
				return "";
			int index = book.getSheetIndex(sheetName);
			
			if (index ==-1)
				return "";
			sheet = book.getSheetAt(index);
			row = sheet.getRow(rowNum-1);
			
			if(row==null)
				return "";
			cell = row.getCell(colNum);
			if(cell==null)
				return "";
					
			if(cell.getCellType()==Cell.CELL_TYPE_STRING)
				return cell.getStringCellValue();
			else if(cell.getCellType()==Cell.CELL_TYPE_NUMERIC || cell.getCellType()==Cell.CELL_TYPE_FORMULA){
				String cellText = String.valueOf(cell.getNumericCellValue());
				if (HSSFDateUtil.isCellDateFormatted(cell)){
					double d = cell.getNumericCellValue();
					Calendar cal = Calendar.getInstance();
					cal.setTime(HSSFDateUtil.getJavaDate(d));
					cellText = (String.valueOf(cal.get(Calendar.YEAR))).substring(2);
					cellText = cal.get(Calendar.MONTH)+1+"/"+
							cal.get(Calendar.DAY_OF_MONTH)+"/"+
							cellText;
				}
				return cellText;
			}else if(cell.getCellType()==Cell.CELL_TYPE_BLANK)
				return "";
			else
				return String.valueOf(cell.getBooleanCellValue());
				
			
			
		} catch(Exception e){
			e.printStackTrace();
			return "Row Number "+rowNum + "or Column Number "+colNum +"doesn't exist in Excel.";
			
		}
		
	}

}

