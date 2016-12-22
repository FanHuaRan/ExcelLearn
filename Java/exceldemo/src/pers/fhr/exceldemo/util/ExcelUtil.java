package pers.fhr.exceldemo.util;

import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import jxl.Cell;
import jxl.CellType;
import jxl.Sheet;
import jxl.Workbook;
/**
 * Excel帮助类
 * @author fhr
 *
 */
public class ExcelUtil {
  /**
   * 以二维字符串数组形式获取Excel
   * @param fileName
   * @return
   */
  public static List<String[][]> readExcel(String fileName){
	  List<String[][]> bookValues=new ArrayList<>();
	  Workbook readWb=null;
		try{
			InputStream inputStream=new FileInputStream(fileName);
			readWb=Workbook.getWorkbook(inputStream);
			 Sheet [] sheets=readWb.getSheets();
			for(Sheet readsheet : sheets){
				int columns=readsheet.getColumns();
				int rows=readsheet.getRows();
				String [][]sheetValue=new String[rows][columns];
				for(int i=0;i<rows;i++){
					for(int j=0;j<columns;j++){
						Cell cell=readsheet.getCell(j, i);
						sheetValue[i][j]=cell.getContents();
					}
				}
				bookValues.add(sheetValue);
			}
			return bookValues;
		}catch(Exception e){
			System.out.print( e.getMessage());
			return null;
		}
  }
  /**
   * 以二维字符串数组形式获取sheet
   * @param fileName
   * @param sheetName
   * @return
   */
  public static String[][] readExcel(String fileName,String sheetName){
	  Workbook readWb=null;
		try{
			InputStream inputStream=new FileInputStream(fileName);
			readWb=Workbook.getWorkbook(inputStream);
			 Sheet sheet=readWb.getSheet(sheetName);
				int columns=sheet.getColumns();
				int rows=sheet.getRows();
				String [][]sheetValue=new String[rows][columns];
				for(int i=0;i<rows;i++){
					for(int j=0;j<columns;j++){
						Cell cell=sheet.getCell(j, i);
						sheetValue[i][j]=cell.getContents();
					}
				}
				return sheetValue;
		}catch(Exception e){
			System.out.print( e.getMessage());
			return null;
		}
  }
  /**
   * 以字典集合形式获取Excel
   * @param fileName
   * @param sheetName
   * @return
   */
  @SuppressWarnings("rawtypes")
public static List<List<Map>> readExcelByMap(String fileName){
	  Workbook readWb=null;
		try{
			InputStream inputStream=new FileInputStream(fileName);
			 readWb=Workbook.getWorkbook(inputStream);
			 List<List<Map>> sheetMaps=new ArrayList<>();
			 Sheet [] sheets=readWb.getSheets();
			 for(Sheet sheet : sheets){
				 List<String> titles=getCollmnTitle(sheet);
				 if(titles==null){
					 return null;
				 }
			    List<Map> rowsMap=new ArrayList<>();
				int columns=sheet.getColumns();
				int rows=sheet.getRows();
				for(int i=1;i<rows;i++){
					Map<String, Object> rowMap=new HashMap<>();
					for(int j=0;j<columns;j++){
						Cell cell=sheet.getCell(j, i);
							rowMap.put(titles.get(j),ParseCellValue(cell));
						}
					rowsMap.add(rowMap);
				}
				sheetMaps.add(rowsMap);
			 }
			return sheetMaps;
		}catch(Exception e){
			System.out.print( e.getMessage());
			return null;
		}
  }
  /**
   * 以字典集合形式获取sheet
   * @param fileName
   * @param sheetName
   * @return
   */
  @SuppressWarnings({ "rawtypes", "rawtypes", "rawtypes", "rawtypes" })
public static List<Map> readExcelByMap(String fileName,String sheetName){
	  Workbook readWb=null;
		try{
			InputStream inputStream=new FileInputStream(fileName);
			 readWb=Workbook.getWorkbook(inputStream);
			 Sheet sheet=readWb.getSheet(sheetName);
			 List<String> titles=getCollmnTitle(sheet);
			 if(titles==null){
				 return null;
			 }
		    List<Map> rowsMap=new ArrayList<>();
			int columns=sheet.getColumns();
			int rows=sheet.getRows();
			for(int i=1;i<rows;i++){
				Map<String, Object> rowMap=new HashMap<>();
				for(int j=0;j<columns;j++){
					Cell cell=sheet.getCell(j, i);
						rowMap.put(titles.get(j),ParseCellValue(cell));
					}
				rowsMap.add(rowMap);
			}
				return rowsMap;
		}catch(Exception e){
			System.out.print( e.getMessage());
			return null;
		}
  }
 @SuppressWarnings({ "rawtypes", "deprecation" })
 private static Object ParseCellValue(Cell cell){
		  Object object=null;
		  String value=cell.getContents();
		  if(value==null||value.equals("")){
			  object=value;
			  return object;
		  }
		  if(cell.getType()==CellType.BOOLEAN){
			  object=Boolean.parseBoolean(value);
		  }
		  else if(cell.getType()==CellType.LABEL){
			  object=value;
		  }
		  else if(cell.getType()==CellType.NUMBER){
			  object=Double.parseDouble(value);
		  }
		  else if(cell.getType()==CellType.DATE){
			  object=Date.parse(value);
		  }
		  else {
			  object=value;
		  }
		  return object;
  }
 public static List<String> getCollmnTitle(Sheet sheet){
		 List<String> columns=new ArrayList<>();
		 if(sheet.getRows()<1||sheet.getColumns()<1){
			 return null;
		 }
		 Cell[] cells=sheet.getColumn(0);
		 for(Cell cell :cells){
			 columns.add(cell.getContents());	
			}
		 return columns;
	}
	  
}
