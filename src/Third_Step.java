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


public class Third_Step {
	 int TableElement=0;
	 int ExcelColumnQuantity=10;
	 int MonoComponentIndicatorColumn=-1;
	 int ComponentColumn=0;
	 String Word;
	 int a=1;
	 int pomniejszacz=0;
	 boolean spr=false;
	 int i=1;

	private  String czytanieliter(int b){
		if(b<Word.length())
		{
			return Character.toString(Word.charAt(b));
		}
		return "";
	}
				
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
	
	public  void porzadkowanie(int ExcelColumnQuantity1, int ComponentColumn1,  int MonoComponentIndicatorColumn1)
	{
		ExcelColumnQuantity=ExcelColumnQuantity1;
		ComponentColumn=ComponentColumn1;
		MonoComponentIndicatorColumn=MonoComponentIndicatorColumn1;	
		struct[] tablica = new struct [ExcelColumnQuantity+7];
		struct[] tablica_nazw = new struct [ExcelColumnQuantity+7];
		struct[] z_soi = new struct[1];
		struct[] koniec = new struct[1];
		z_soi[0]=new struct("",-1);
		koniec[0]=new struct("",-1);
		
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
			//tworzy nowy arkusz o nazwie "przed rozprocentowaniem
			        XSSFSheet sheet1 = workbook.createSheet("ostateczne");
			//tworzy iterator dla arkusza zerowego
			        java.util.Iterator<Row> rowIterator = sheet.iterator();
			        Row row = rowIterator.next();
			        System.out.println("Third step started");
			        for(TableElement=0; TableElement<ExcelColumnQuantity+6; TableElement++)
					{ 
						czytaniezExcella(TableElement,row, tablica_nazw);
					}
			        TableElement=0;
			        
			        while(rowIterator.hasNext()) 
			        {
						row = rowIterator.next();
						i++;
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
						Word=tablica[ComponentColumn].text;

						
					if(rowIterator.hasNext())
					{
						Row row1 = sheet.getRow(i);
						Cell cell = row1.getCell(ComponentColumn);
						Cell cell2 = row1.getCell(ExcelColumnQuantity+5);
						
						if (cell!=null)
							{	
								z_soi[0].text=cell.getStringCellValue();		
						
							}	
						 if (cell2!=null)
							{	
								koniec[0].liczba= cell2.getNumericCellValue();
							}
						}
						
					
					if(tablica[ComponentColumn].text.equals("lecytyny")||tablica[ComponentColumn].text.equals("lecytyny "))
					{
						if(z_soi[0].text.equals("z soi")|| z_soi[0].text.equals("z SOI") || z_soi[0].text.equals("soja")|| z_soi[0].text.equals("SOJA"))
						{
							tablica[ComponentColumn].text="lecytyny z soi";
							tablica[MonoComponentIndicatorColumn].liczba=0;	
							tablica[ExcelColumnQuantity+5].liczba+=koniec[0].liczba;
		
							spr=true;
						}
					}
					
					if((tablica[ComponentColumn].text.equals("z soi") || tablica[ComponentColumn].text.equals("z SOI") || tablica[ComponentColumn].text.equals("soja")|| tablica[ComponentColumn].text.equals("SOJA"))&& spr)
					{	
					spr=false;
					pomniejszacz++;
					}
					else
					{
						
						Row header = sheet1.createRow(a);
						a++;
						if(tablica[ExcelColumnQuantity].liczba>1)
						{
							tablica[ExcelColumnQuantity].liczba-=pomniejszacz;
						}
						else if(tablica[ExcelColumnQuantity].liczba==1)
						{
							pomniejszacz=0;
						}
						for(TableElement=0; TableElement<ExcelColumnQuantity+6; TableElement++)
						{
							zapisDoExcella(header, TableElement, tablica);
						}
						TableElement=0;		
					}	
			
			        }
        
			        Row header = sheet1.createRow(0);
			        System.out.println("Third step is saving its progress...");
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
					   			System.out.println("no file exception");
							} 
					   	catch (IOException e) 
					   		{
					   			e.printStackTrace();
					   			System.out.println("error");
					   		}
			        
			        }

		}
		
	
	
	
	
	
	
	
	
	
	
	
