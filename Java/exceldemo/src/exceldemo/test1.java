

package exceldemo;

import java.io.FileInputStream;
import java.io.InputStream;

import jxl.Cell;
import jxl.CellType;
import jxl.Sheet;
import jxl.Workbook;

public class test1{
	public  static void main(String[]args){
		Workbook readWb=null;
		try{
			InputStream inputStream=new FileInputStream("（徐发）11111.xls");
			readWb=Workbook.getWorkbook(inputStream);
			 Sheet [] sheets=readWb.getSheets();
			for(Sheet readsheet : sheets){
				int columns=readsheet.getColumns();
				int rows=readsheet.getRows();
				for(int i=0;i<rows;i++){
					for(int j=0;j<columns;j++){
						Cell cell=readsheet.getCell(j, i);
						System.out.print(cell.getContents()+" ");
					}
					System.out.print("\n");
				}
			}
		}catch(Exception e){
			System.out.print( e.getMessage());
		}
	}
}
