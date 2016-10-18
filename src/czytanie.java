import java.io.*;
import java.math.BigInteger;

import javax.swing.text.html.HTMLDocument.Iterator;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.*;

import com.microsoft.schemas.office.visio.x2012.main.CellType;

public class czytanie 
{
	static String slowo=null;
	static String pociete="";
	static String c = null;
	static String d=null;
	static String a=null;
	static String e=null;
	static String fi="";
	static int i=0;
	
	static boolean czyuzyteproductitem=false;
	static double Product_item;
	static double Pomocnicza_product_item=75289;
	static double Wieloskladnik;
	static double id_Skladnik_algorytm;
	static double id_bb_product;
	static double EAN;
	static double id_skladnik;
	static double Pozycja_skladnika;
	static double Id_Parent;
	
	static String Skladnik_algorytm;
	static String Skladnik;
	static String Pelen_sklad;
	static String Nazwa_produktu;
	static String Grupa_BB;

	static String ktorypociety;
	static double iddosprawdzenia=1;
	static int zmienna=0;
	static int zmienna1=0;
	static double pomocniczadoparent;
	static String f="";
	
	private static boolean IfEqualsNumber(String c){
		if (c.equals("0") || c.equals("1") || c.equals("2") || 
			c.equals("3") || c.equals("4") || c.equals("5") ||
			c.equals("6") || c.equals("7") || c.equals("8") || 
			c.equals("9"))
			{
				return true;
			}
		return false;
	}
	
	private static String czytanieZExcellaStringa(int a, Row row){
		Cell cell = row.getCell(a);
		if (cell == null)
		{
			return "";
		}
		else
		{
			switch(cell.getCellType()) {
			
				case Cell.CELL_TYPE_NUMERIC:
					double wartosc01=cell.getNumericCellValue();
					long x = (long) wartosc01;
					return Long.toString(x);
			
				case Cell.CELL_TYPE_STRING:
					return cell.getStringCellValue();
			
				case Cell.CELL_TYPE_BLANK:
					return "";
			}
		}
		return null;
	}

	private static double czytanieZExcellaIntegera(int a, Row row){
		Cell cell = row.getCell(a);
		if (cell == null)
		{
			return -1;
		}
		else
		{
		switch(cell.getCellType()) {
		
		case Cell.CELL_TYPE_NUMERIC:
			double wartosc01=cell.getNumericCellValue();
			long x = (long) wartosc01;
			return x;
	
		case Cell.CELL_TYPE_BLANK:
			return -1;
			}
		}
		return -1;
	}

	private static void zapisDoExcella(int a, String wartosc, double wartosc01, Row row){
		if (wartosc != "" && wartosc01 == -1)
			row.createCell(a).setCellValue(wartosc);
		else if (wartosc == "" && wartosc01 != -1)
			row.createCell(a).setCellValue(wartosc01);
		else
			row.createCell(a).setCellType(Cell.CELL_TYPE_BLANK);	
	}
	
   public static void main(String[] args)throws Exception 
   {
	   try {
			
			FileInputStream file = new FileInputStream(new File("D:\\test.xlsx"));
//Get the workbook instance for XLS file 
			XSSFWorkbook workbook = new XSSFWorkbook(file);
//Get first sheet from the workbook
			XSSFSheet sheet = workbook.getSheetAt(0);
			XSSFSheet sheet1 = workbook.createSheet("Przed rozprocentowaniem");
//Create a new row in current sheet
			
			java.util.Iterator<Row> rowIterator = sheet.iterator();
		
			while(rowIterator.hasNext()) {
				Row row = rowIterator.next();

				Wieloskladnik			=	czytanieZExcellaIntegera(0, row);
				Product_item			=	czytanieZExcellaIntegera(1, row);
				id_Skladnik_algorytm	=	czytanieZExcellaIntegera(2, row);
				Skladnik_algorytm		=	czytanieZExcellaStringa	(3, row);
				Skladnik				=	czytanieZExcellaStringa	(4, row);
				id_skladnik				=	czytanieZExcellaIntegera(5, row);
				Pozycja_skladnika		=	czytanieZExcellaIntegera(6, row);
				id_bb_product			=	czytanieZExcellaIntegera(7, row);
				EAN						=	czytanieZExcellaIntegera(8, row);
				Nazwa_produktu			=	czytanieZExcellaStringa	(9, row);
				Grupa_BB				=	czytanieZExcellaStringa	(10, row);
				
				czyuzyteproductitem=false;

//Pozycja skladnika i ponizsze funkcje pomagaja ustalic pozycje skladnika w skladzie
					if (Pozycja_skladnika!=1)
					{
						Pozycja_skladnika=iddosprawdzenia;
					}
					else
					{
						iddosprawdzenia=1;
					}
//jezeli aktualnie badany skladnik to wieloskladnik to go trzeba rozmontowac					
					if(Wieloskladnik==1)
					{	
						int numerek=0;
						Id_Parent=Product_item;
						slowo=Skladnik;
						pomocniczadoparent=Id_Parent;
						
						for(int b=0;b<=slowo.length();b++)
						{
// 192-208 CZYTANIE KOLEJNYCH LITER POBRANEGO SKLADU
							if(b+1<=slowo.length())
							{
								a = Character.toString(slowo.charAt(b));
							}
							if(b+2<slowo.length())
							{
								c = Character.toString(slowo.charAt(b+1));
							}
							if(b+3<slowo.length())
							{
								d = Character.toString(slowo.charAt(b+2));
							}
							if(b+4<slowo.length())
							{
								e = Character.toString(slowo.charAt(b+4));
							}
							if(b+5<slowo.length())
							{
								f = Character.toString(slowo.charAt(b+5));
							}
							if(b-1>=0)
							{
								fi = Character.toString(slowo.charAt(b-1));
							}
						
//b- aktualny numer litery na ktorej stoimy
//pociete- jeden element skladu, ktory jest samodzielny
//i chcemy go wyszczegolnic jako pojedynczy
							if (b+1>slowo.length())
							{
//jezeli to koniec stringa zapisujemy go od razu zeby nie powtarzac
//dwa razy ostatniej litery
								if(pociete!="")
								{
	
									Row header = sheet1.createRow(i);
									if (czyuzyteproductitem)
									{
									Product_item=Pomocnicza_product_item;
									Pomocnicza_product_item++;
									}
									else
									{
										czyuzyteproductitem=true;
									}
							
									zapisDoExcella(0,	 "",				Wieloskladnik,			 header);						
									zapisDoExcella(1,	 "",				Product_item,			 header);						
									zapisDoExcella(2,	 "",				id_Skladnik_algorytm,	 header);						
									zapisDoExcella(3,	 Skladnik_algorytm,	-1,						 header);	
									zapisDoExcella(4,	 pociete,			-1,						 header);
									zapisDoExcella(5,	 "", 				id_skladnik,			 header);
									zapisDoExcella(6,	 "",				Pozycja_skladnika,		 header);
									zapisDoExcella(7,	 "", 				id_bb_product,			 header);
									zapisDoExcella(8,	 "", 				EAN,					 header);
									zapisDoExcella(9,	 Nazwa_produktu,	-1,						 header);
									zapisDoExcella(10,	 Grupa_BB, 			-1,						 header);
									zapisDoExcella(11,	 "", 				Id_Parent,				 header);
									zapisDoExcella(12,	 "", 				-1,						 header);
									zapisDoExcella(13,	 Pelen_sklad,		-1,						 header);
									Wieloskladnik=0;
									pociete="";
									i++;
									Pozycja_skladnika++;
								}
							}
// 261-275 sprawdzamy czy sa krzaki w postaci <sdasdasa>
// jezeli tak omija je i b przypisuje numer litery pod ktora jest >
							else if(a.equals("<"))
							{
								for(int iterator=b;iterator<slowo.length();iterator++)
								{
									a = Character.toString(slowo.charAt(iterator));
									if(a.equals(">"))
									{
										b=iterator;
										break;
									}
								}
							}

//patrzy czy to koniec pojedynczego skladnika, jezeli tak to zapisze go jak nie to leci dalej
							else if ((a.equals(",") && !IfEqualsNumber(c)) ||
									(a.equals(")") && !IfEqualsNumber(d)) || 
									(a.equals("(") && !IfEqualsNumber(c)) || 
									 a.equals(":") || a.equals("[") ||
									 a.equals("]") || a.equals("/")  || 
									 a.equals(";") || b+1>slowo.length() ||
									 (a.equals("i") && fi.equals(" ") && c.equals(" ")))
							{
								if(pociete!="")
								{
									Row header = sheet1.createRow(i);

									if (czyuzyteproductitem)
									{
									
									Product_item=Pomocnicza_product_item;
									Pomocnicza_product_item++;
									}
									else
									{
										czyuzyteproductitem=true;
									}
									if(Id_Parent==Product_item)
									{
										zapisDoExcella(11,	 "", 				0,				 header);
									}
									else
										zapisDoExcella(11,	 "", 				Id_Parent,				 header);
									
									if (a.equals("("))
									{
										Id_Parent=Product_item;
										Wieloskladnik=1;
									}
										
									zapisDoExcella(0,	 "",				Wieloskladnik,			 header);						
									zapisDoExcella(1,	 "",				Product_item,			 header);						
									zapisDoExcella(2,	 "",				id_Skladnik_algorytm,	 header);						
									zapisDoExcella(3,	 Skladnik_algorytm,	-1,						 header);	
									zapisDoExcella(4,	 pociete,			-1,						 header);
									zapisDoExcella(5,	 "", 				id_skladnik,			 header);
									zapisDoExcella(6,	 "",				Pozycja_skladnika,		 header);
									zapisDoExcella(7,	 "", 				id_bb_product,			 header);
									zapisDoExcella(8,	 "", 				EAN,					 header);
									zapisDoExcella(9,	 Nazwa_produktu,	-1,						 header);
									zapisDoExcella(10,	 Grupa_BB, 			-1,						 header);
									//zapisDoExcella(11,	 "", 				Id_Parent,				 header);
									zapisDoExcella(12,	 "", 				-1,						 header);
									
									
								
									
									id_skladnik=-1;
								
									Wieloskladnik=0;
									pociete="";
									++Pozycja_skladnika;
									i++;
									
									if( zmienna1==0)
									{
										header.createCell(13).setCellValue(Skladnik);
										zmienna1++;
									}
									else
									{
										header.createCell(13).setCellType(Cell.CELL_TYPE_BLANK);
									}
						
									if (a.equals(")"))
									{
										Id_Parent=pomocniczadoparent;
									}
									numerek++;

									pociete="";
								}
							}
// nie chcemy spacji na poczatku skladnika
							else if (a.equals(" ") && pociete.equals(""))
							{
//pusto, omija spacje							
							}
//jezeli to normalna litera dokleja ja do przyszlego skladnika i od poczatku for
							else
							{
								pociete=pociete+a;
							}
					
						}
//wieloskladnik rozmontowany ustawiamy pomocnicze i lecimy dalej
						Id_Parent=0;
						zmienna1=0;
						zmienna=0;
						iddosprawdzenia=Pozycja_skladnika;
					}
//monoskladnik		
					else
					{
						slowo=Skladnik;
// dla monoskladnikow czysci z syfu <sabduagd>
						for(int b=0;b<=slowo.length();b++)
						{
							if(b+1<=slowo.length())
							{
								a = Character.toString(slowo.charAt(b));
							}
							else
							{
								break;
							}
							if(a.equals("<"))
							{
								for(int iterator=b;iterator<slowo.length();iterator++)
								{
									a = Character.toString(slowo.charAt(iterator));
									if(a.equals(">"))
									{
										b=iterator+1;
										a = Character.toString(slowo.charAt(b));
										break;
									}
								}
							}
						pociete=pociete+a;
						}
					Skladnik=pociete;
					pociete="";
// ZAPIS DO PROGRAMU JEZELI NIE JEST WIELOSKLADNIKIEM	
					Row header = sheet1.createRow(i);
					zapisDoExcella(0,	 "",				Wieloskladnik,			 header);						
					zapisDoExcella(1,	 "",				Product_item,			 header);						
					zapisDoExcella(2,	 "",				id_Skladnik_algorytm,	 header);						
					zapisDoExcella(3,	 Skladnik_algorytm,	-1,						 header);	
					zapisDoExcella(4,	 Skladnik,			-1,						 header);
					zapisDoExcella(5,	 "", 				id_skladnik,			 header);
					zapisDoExcella(6,	 "",				Pozycja_skladnika,		 header);
					zapisDoExcella(7,	 "", 				id_bb_product,			 header);
					zapisDoExcella(8,	 "", 				EAN,					 header);
					zapisDoExcella(9,	 Nazwa_produktu,	-1,						 header);
					zapisDoExcella(10,	 Grupa_BB, 			-1,						 header);
					zapisDoExcella(11,	 "", 				Id_Parent,				 header);
					zapisDoExcella(12,	 "", 				-1,						 header);
					zapisDoExcella(13,	 Pelen_sklad,		-1,						 header);
					id_skladnik=-1;
					i++;
					iddosprawdzenia++;
					}
						
				}
					
//koniec programu	
				Row header = sheet1.createRow(0);
				header.createCell(0).setCellValue("Wieloskladnik");
				header.createCell(1).setCellValue("Product_item");
				header.createCell(2).setCellValue("id_Skladnik_algorytm");
				header.createCell(3).setCellValue("Skladnik_algorytm");
				header.createCell(4).setCellValue("Skladnik");
				header.createCell(5).setCellValue("id_skladnik");
				header.createCell(6).setCellValue("Pozycja_skladnika");
				header.createCell(7).setCellValue("id_bb_product");
				header.createCell(8).setCellValue("EAN");
				header.createCell(9).setCellValue("Nazwa_produktu");
				header.createCell(10).setCellValue("Grupa BB");
				header.createCell(11).setCellValue("Id_Parent");
				header.createCell(12).setCellValue("<puste potrzebne potem>");
				header.createCell(13).setCellValue("pelen sklad wieloskladnika");
				
				int index = workbook.getSheetIndex(sheet);
				workbook.removeSheetAt(index);
				
				file.close();
				FileOutputStream out = new FileOutputStream(new File("D:\\test.xlsx"));
				workbook.write(out);
				out.close();
				System.out.println("Zrobiona pierwsza czesc skryptu nie otwieraj jeszcze bazy");
			
			} 
	   		catch (FileNotFoundException e) 
	   		{
				e.printStackTrace();
	   		} 
	   		catch (IOException e) 
	   		{
	   			e.printStackTrace();
	   		}
   		}
	}


