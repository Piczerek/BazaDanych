import java.io.*;
import java.math.BigInteger;
import java.util.Scanner;

import javax.swing.text.html.HTMLDocument.Iterator;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.*;

import com.microsoft.schemas.office.visio.x2012.main.CellType;


public class Second_Step {

	
	
	 int TableElement=0;
	 int ExcelColumnQuantity=10;
	 int MonoComponentIndicatorColumn=-1;
	 int ComponentColumn=0;
	 int Row_to_save_in=1;
	 String CutComponent="";
	 String Percent="";

		
		
	private  boolean IfEqualsNumber(String c){
		if (c.equals("0") || c.equals("1") || c.equals("2") || 
			c.equals("3") || c.equals("4") || c.equals("5") ||
			c.equals("6") || c.equals("7") || c.equals("8") || 
			c.equals("9"))
			{
				return true;
			}
		return false;
	}
		
	private int FindPercent(String word){
		int i;
		for(i=0;i<word.length();i++)
		{
			if(Character.toString(word.charAt(i)).equals("%"))
			{
				return i;
			}
		}
		return -1;
	}
	
	private void CutPercent(String word, int PercentPosition){
		int SpaceBeforePercent=0;
		
		for(int a=PercentPosition-1;a>=0;a--)
		{
			if(a==0)
			{
				SpaceBeforePercent=a;
				break;
			}
			if( !(IfEqualsNumber(Character.toString(word.charAt(a))) || Character.toString(word.charAt(a)).equals(",")) && !Character.toString(word.charAt(a+1)).equals("%"))
			{
				SpaceBeforePercent=a;
				break;
			}
			
		}
		for(int b=0;b<SpaceBeforePercent;b++)
		{
			CutComponent+=Character.toString(word.charAt(b));
		}
		for(int b=PercentPosition+1;b<word.length();b++)
		{
			CutComponent+=Character.toString(word.charAt(b));
		}
		for(int b=SpaceBeforePercent+1;b<=PercentPosition;b++)
		{
			Percent+=Character.toString(word.charAt(b));
		}
		
	}
	
	
	private  void zapisDoExcella(Row row, int i, struct tablica[]){
		if (tablica[i].text != "" && tablica[i].liczba == -1)
			row.createCell(i).setCellValue(tablica[i].text);
		else if (tablica[i].text == "" && tablica[i].liczba != -1)
			row.createCell(i).setCellValue(tablica[i].liczba);
		else
			row.createCell(i).setCellType(Cell.CELL_TYPE_BLANK);	
	}	
		
	private  void czytaniezExcella(int a, Row row, struct tablica[]){
		Cell cell = row.getCell(a);
	 if (cell!=null)
		{
			switch(cell.getCellType()) {		
				case Cell.CELL_TYPE_NUMERIC:
					tablica[TableElement].liczba= cell.getNumericCellValue();
					break;
				case Cell.CELL_TYPE_STRING:	
					tablica[TableElement].text=cell.getStringCellValue();
					break;
			}
		}	
	}
	
	
	
	
	public void porzadkowanie(int ExcelColumnQuantity1, int ComponentColumn1,  int MonoComponentIndicatorColumn1)
	{
		this.ExcelColumnQuantity=ExcelColumnQuantity1;
		this.ComponentColumn=ComponentColumn1;
		this.MonoComponentIndicatorColumn=MonoComponentIndicatorColumn1;
	
		struct[] tablica = new struct [ExcelColumnQuantity+7];
		struct[] tablica_nazw = new struct [ExcelColumnQuantity+7];
		for(TableElement=0; TableElement<ExcelColumnQuantity+6; TableElement++)
		{ 
			tablica[TableElement]=new struct("",-1);
		}
		TableElement=0;
		for(TableElement=0; TableElement<ExcelColumnQuantity+6; TableElement++)
		{ 
			tablica_nazw[TableElement]=new struct("",-1);
		}
		TableElement=0;
		
	     
	
		try {		  
			// tworzy input stream 
					InputStream is = new FileInputStream("D:\\eclipse\\BazaDanych-master\\tescik.xlsx");
			//sczytuje excela
					XSSFWorkbook workbook = new XSSFWorkbook(is);
			//bierze pierwszy arkusz
			        XSSFSheet sheet = workbook.getSheetAt(0);
			//tworzy nowy arkusz o nazwie "przed rozPercentowaniem
			        XSSFSheet sheet1 = workbook.createSheet("rozPercentowane");
			//tworzy iterator dla arkusza zerowego
			        java.util.Iterator<Row> rowIterator = sheet.iterator();
			        Row row = rowIterator.next();
			        
			        System.out.println("Second step started");
			        
			        for(TableElement=0; TableElement<ExcelColumnQuantity+6; TableElement++)
					{ 
						czytaniezExcella(TableElement,row, tablica_nazw);
					}
				TableElement=0;
				
				
			        while(rowIterator.hasNext()) 
			        {
						row = rowIterator.next();
			//czyœci wszystkie oczka tablicy po kolei wierszy
						for(TableElement=0; TableElement<ExcelColumnQuantity+6; TableElement++)
						{ 
							tablica[TableElement].text="";
							tablica[TableElement].liczba=-1;
						}
					TableElement=0;
					
			// czyta po kolei kolumny w danym rzedzie			
						for(TableElement=0; TableElement<ExcelColumnQuantity+6; TableElement++)
							{ 
								czytaniezExcella(TableElement,row, tablica);
							}
						TableElement=0;
				
						
						if (tablica[MonoComponentIndicatorColumn].liczba==0)
						{
							int a=FindPercent(tablica[ComponentColumn].text);
							if(a!=-1)
							{
								CutPercent(tablica[ComponentColumn].text, a);
							}
							else
							{
								CutComponent=tablica[ComponentColumn].text;
								Percent="";
							}
								
						tablica[ExcelColumnQuantity+3].text=Percent;
						tablica[ComponentColumn].text=CutComponent;
						}
						CutComponent="";
						Percent="";
						
						Row header = sheet1.createRow(Row_to_save_in);
						Row_to_save_in++;
						for(TableElement=0; TableElement<ExcelColumnQuantity+6; TableElement++)
						{
							zapisDoExcella(header, TableElement, tablica);
						}
						TableElement=0;
		
			        }
			        Row header = sheet1.createRow(0);
			        tablica_nazw[ExcelColumnQuantity+3].text="zawartosc Percentowa";
			        System.out.println("second step is saving its progress...");
			    	for(TableElement=0; TableElement<ExcelColumnQuantity+6; TableElement++)
			    		{
			    			zapisDoExcella(header, TableElement, tablica_nazw);
			    		}
			    	TableElement=0;
			    	
			        
						int index = workbook.getSheetIndex(sheet);
						workbook.removeSheetAt(index);			
						is.close();
						FileOutputStream out = new FileOutputStream(new File("D:/eclipse/BazaDanych-master/tescik.xlsx"));
						workbook.write(out);
						out.close();
			        
			}
			        	
						
					   	catch (FileNotFoundException e) 
							{
					   			e.printStackTrace();
					   			System.out.println("file not found");
							} 
					   	catch (IOException e) 
					   		{
					   			e.printStackTrace();
					   			System.out.println("error");
					   		}
			        
			        }

		
		
	
	
	
	
	
	
	
	
	
	
	
}