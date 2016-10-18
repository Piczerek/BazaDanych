import java.io.*;
import java.math.BigInteger;

import javax.swing.text.html.HTMLDocument.Iterator;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.*;

import com.microsoft.schemas.office.visio.x2012.main.CellType;

public class czytanieDaryNatury
{
	static String slowo=null;
	static String pociete="";
	static String c = null;
	static String d=null;
	static String a=null;
	static String e=null;
	static String fi="";
	static int i=0;
	static double Id_Parent;
	static double Pozycja_skladnika=0;
	static double Wieloskladnik;
	
	static boolean czyuzyteproductitem=false;
	static String ID;
	static String nazwa_artykulu;
	static String zwartosc_netto;
	static String waga_brutto;
	static String skladniki;
	static String bezglutenowe;
	static String opis_produktu;
	static String sposob_uzycia;
	static String kraj_pochodzenia;
	static String angnazwa;
	static String hgw;
	


	
	static double ilosc_sztuk_w_zbiorczym;
	static double ean;
	static double wysokosckartonika;
	static double dlugosckartonika;
	static double szerokosckartonika;
	static double dlugosckartzbior;
	static double wyoskosckartzbior;
	static double szerokosckartzbior;
	static double iloscsztukwwarstwie;
	static double iloscwarstwnapalecie;
	static double iloscsztuknapalecie;
	
	
	static String ktorypociety;
	static double iddosprawdzenia=1;
	static int zmienna=0;
	static int zmienna1=0;
	static double pomocniczadoparent;
	static String f="";
	
	
	
	public static String sprawdzenieczyjestsyf(String slowodowyczyszczenia){
		String sprawdzacz="";
		String literka = "";
		String wyczyszczone="";
		String wyczyszczon2="";
		if(slowodowyczyszczenia.equals("*produkt rolnictwa ekologicznego") || slowodowyczyszczenia.equals("produkt rolnictwa ekologicznego") ||slowodowyczyszczenia.equals(" *produkt rolnictwa ekologicznego"))
		{
			wyczyszczone="";
		}
		else if(slowodowyczyszczenia.equals("w ró¿nych proporcjach") || slowodowyczyszczenia.equals(" w ró¿nych proporcjach") || slowodowyczyszczenia.equals("w ró¿nych proporcjach "))
		{
			wyczyszczone="";
		}
		else
		{ wyczyszczone="";
			for (int costam=0;costam<=slowodowyczyszczenia.length();costam++) 
			{
				for (int costam2=costam;costam2<=(costam+30);costam2++)
				{
					if(costam2+1<=slowodowyczyszczenia.length())
					{
						literka = Character.toString(slowodowyczyszczenia.charAt(costam2));
						sprawdzacz=sprawdzacz+literka;
					}
					
				}
				System.out.println(sprawdzacz);
				if (sprawdzacz.equals("produkt rolnictwa ekologicznego"))
					costam=costam+31;
				sprawdzacz="";
				if(costam+1<=slowodowyczyszczenia.length())
				{
				literka= Character.toString(slowodowyczyszczenia.charAt(costam));
				
				wyczyszczone=wyczyszczone+literka;
			}}
			for (int costam=0;costam<=wyczyszczone.length();costam++) 
			{
				for (int costam2=costam;costam2<=(costam+20);costam2++)
				{
					if(costam2+1<=wyczyszczone.length())
					{
						literka = Character.toString(wyczyszczone.charAt(costam2));
						sprawdzacz=sprawdzacz+literka;
					}
					
				}
				System.out.println(sprawdzacz);
				if (sprawdzacz.equals("w ró¿nych proporcjach"))
					costam=costam+21;
				sprawdzacz="";
				if(costam+1<=wyczyszczone.length())
				{
				literka= Character.toString(wyczyszczone.charAt(costam));
				
				wyczyszczon2=wyczyszczon2+literka;
			}}
			
		}
		
		
	return wyczyszczon2;
	}
	
	
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
					
				case Cell.CELL_TYPE_FORMULA:
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
			return wartosc01;
	
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

				ID						=	czytanieZExcellaStringa(0, row);
				nazwa_artykulu			=	czytanieZExcellaStringa(1, row);
				//zwartosc_netto			=	czytanieZExcellaStringa(2, row);
				waga_brutto				=	czytanieZExcellaStringa	(3, row);
				ilosc_sztuk_w_zbiorczym	=	czytanieZExcellaIntegera(4, row);
				ean						=	czytanieZExcellaIntegera(5, row);
				skladniki				=	czytanieZExcellaStringa(6, row);
				bezglutenowe			=	czytanieZExcellaStringa(7, row);
				opis_produktu			=	czytanieZExcellaStringa(8, row);
				sposob_uzycia			=	czytanieZExcellaStringa	(9, row);
				kraj_pochodzenia		=	czytanieZExcellaStringa	(10, row);
				wysokosckartonika		=	czytanieZExcellaIntegera(11, row);
				dlugosckartonika		=	czytanieZExcellaIntegera(12, row);
				szerokosckartonika		=	czytanieZExcellaIntegera(13, row);
				dlugosckartzbior		=	czytanieZExcellaIntegera(14, row);
				szerokosckartzbior		=	czytanieZExcellaIntegera(15, row);
				wyoskosckartzbior		=	czytanieZExcellaIntegera(16, row);
				iloscsztukwwarstwie		=	czytanieZExcellaIntegera(17, row);
				iloscwarstwnapalecie	=	czytanieZExcellaIntegera(18, row);
				iloscsztuknapalecie		=	czytanieZExcellaIntegera(19, row);
				angnazwa				=	czytanieZExcellaStringa(20, row);
				hgw						=	czytanieZExcellaStringa(21, row);


				
//jezeli aktualnie badany skladnik to wieloskladnik to go trzeba rozmontowac					
					if(true)
					{	
						int numerek=0;
						Id_Parent=pomocniczadoparent;
						slowo=skladniki;
						
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
									String zmienna=sprawdzenieczyjestsyf(pociete);
									pociete=zmienna;
									if (pociete!="")
									{
									Row header = sheet1.createRow(i);
									Wieloskladnik=0;
									zapisDoExcella(0,	 ID,				-1,							 header);
									zapisDoExcella(1,	 nazwa_artykulu,	-1,	 						 header);
									zapisDoExcella(2,	 zwartosc_netto,	-1,							 header);
									zapisDoExcella(3,	 waga_brutto,		-1,							 header);
									zapisDoExcella(4,	 "",				ilosc_sztuk_w_zbiorczym,	 header);
									zapisDoExcella(5,	 "",				ean,	 					 header);
									zapisDoExcella(6,	 pociete,			-1,							 header);
									zapisDoExcella(7,	 bezglutenowe,		-1,							 header);
									zapisDoExcella(8,	 opis_produktu,		-1,							 header);
									zapisDoExcella(9,	 sposob_uzycia,		-1,							 header);
									zapisDoExcella(10,	 kraj_pochodzenia,	-1,							 header);
									zapisDoExcella(11,	 "",				wysokosckartonika,			 header);
									zapisDoExcella(12,	 "",				dlugosckartonika,			 header);
									zapisDoExcella(13,	 "",				szerokosckartonika,			 header);
									zapisDoExcella(14,	 "",				dlugosckartzbior,			 header);
									zapisDoExcella(15,	 "",				szerokosckartzbior,			 header);
									zapisDoExcella(16,	 "",				wyoskosckartzbior,			 header);
									zapisDoExcella(17,	 "",				iloscsztukwwarstwie,		 header);
									zapisDoExcella(18,	 "",				iloscwarstwnapalecie,		 header);
									zapisDoExcella(19,	 "",			iloscsztuknapalecie,	 						 header);
									zapisDoExcella(20,	 angnazwa,			-1,	 						 header);
									zapisDoExcella(21,	 hgw,				-1,							 header);
									zapisDoExcella(22,	 "",				Wieloskladnik,				 header);
									if (Wieloskladnik==1)
									{
									zapisDoExcella(23,	 "",		0,						 header);
										
									}
									else
									zapisDoExcella(23,	 "",		Id_Parent,						 header);
									
									zapisDoExcella(24,	 "",		Pozycja_skladnika,						 header);
									
									Wieloskladnik=0;
									pociete="";
									i++;
									Pozycja_skladnika++;
								}}
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
									 (a.equals("i") && fi.equals(" ") && c.equals(" ")) ||
									 (a.equals("o") && fi.equals(" ") && c.equals("r") && d.equals("a") && e.equals(" ")) )
							{
								if(pociete!="")
								{
									if (a.equals("o") && fi.equals(" ") && c.equals("r") && d.equals("a") && e.equals(" "))
									{
										b=b+3;
									}
									String zmienna=sprawdzenieczyjestsyf(pociete);
									pociete=zmienna;
									if (pociete!="")
									{
									
									Row header = sheet1.createRow(i);

									
									if (a.equals("("))
									{
										Id_Parent=Pozycja_skladnika;
										Wieloskladnik=1;
									}
										Wieloskladnik=0;
									zapisDoExcella(0,	 ID,				-1,							 header);
									zapisDoExcella(1,	 nazwa_artykulu,	-1,	 						 header);
									zapisDoExcella(2,	 zwartosc_netto,	-1,							 header);
									zapisDoExcella(3,	 waga_brutto,		-1,							 header);
									zapisDoExcella(4,	 "",				ilosc_sztuk_w_zbiorczym,	 header);
									zapisDoExcella(5,	 "",				ean,	 					 header);
									zapisDoExcella(6,	 pociete,			-1,							 header);//////
									zapisDoExcella(7,	 bezglutenowe,		-1,							 header);
									zapisDoExcella(8,	 opis_produktu,		-1,							 header);
									zapisDoExcella(9,	 sposob_uzycia,		-1,							 header);
									zapisDoExcella(10,	 kraj_pochodzenia,	-1,							 header);
									zapisDoExcella(11,	 "",				wysokosckartonika,			 header);
									zapisDoExcella(12,	 "",				dlugosckartonika,			 header);
									zapisDoExcella(13,	 "",				szerokosckartonika,			 header);
									zapisDoExcella(14,	 "",				dlugosckartzbior,			 header);
									zapisDoExcella(15,	 "",				szerokosckartzbior,			 header);
									zapisDoExcella(16,	 "",				wyoskosckartzbior,			 header);
									zapisDoExcella(17,	 "",				iloscsztukwwarstwie,		 header);
									zapisDoExcella(18,	 "",				iloscwarstwnapalecie,		 header);
									zapisDoExcella(19,	 "",			iloscsztuknapalecie,	 						 header);
									zapisDoExcella(20,	 angnazwa,			-1,	 						 header);
									zapisDoExcella(21,	 hgw,				-1,							 header);
									zapisDoExcella(22,	 "",				Wieloskladnik,				 header);
									
									//zapisDoExcella(22,	 skladniki,		-1,						 header);
									if (Wieloskladnik==1)
									{
									zapisDoExcella(23,	 "",		0,						 header);
										
									}
									else
									zapisDoExcella(23,	 "",		Id_Parent,						 header);
									
									zapisDoExcella(24,	 "",		Pozycja_skladnika,						 header);
									
								
									
									

									pociete="";
									++Pozycja_skladnika;
									i++;
									Wieloskladnik=0;
									//////////////////////////////////////////////////////////////zmienic nr dla pelnego skladu
									if( zmienna1==0)
									{
										header.createCell(25).setCellValue(skladniki);
										zmienna1++;
									}
									else
									{
										header.createCell(25).setCellType(Cell.CELL_TYPE_BLANK);
									}
						
									if (a.equals(")"))
									{
										Id_Parent=pomocniczadoparent;
									}
									numerek++;

									pociete="";
								}}
							}
// nie chcemy spacji na poczatku skladnika
							else if ((a.equals(" ") || a.equals(".")) && pociete.equals(""))
							{
//pusto, omija spacje	lub kropke
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

						
				}
					
//koniec programu	
				Row header = sheet1.createRow(0);
				header.createCell(0).setCellValue("ID");
				header.createCell(1).setCellValue("nazwa artykulu");
				header.createCell(2).setCellValue("zawartosc netto");
				header.createCell(3).setCellValue("waga brutto");
				header.createCell(4).setCellValue("ilosc sztuk w zbiorczym");
				header.createCell(5).setCellValue("ean");
				header.createCell(6).setCellValue("skladniki");
				header.createCell(7).setCellValue("bezglutenowe");
				header.createCell(8).setCellValue("opis produktu");
				header.createCell(9).setCellValue("sposob uzycia");
				header.createCell(10).setCellValue("kraj pochodzenia");
				header.createCell(11).setCellValue("wysokosc kartonika");
				header.createCell(12).setCellValue("dlugosc kartonika");
				header.createCell(13).setCellValue("szerokosc kartonika");
				header.createCell(14).setCellValue("dlugosc kartonika zbior");
				header.createCell(15).setCellValue("szerokosc kartonika zbior");
				header.createCell(16).setCellValue("wysokosc kartonika zbior");
				header.createCell(17).setCellValue("ilosc sztuk w warstwie");
				header.createCell(18).setCellValue("ilosc warstw na palecie");
				header.createCell(19).setCellValue("ilosc sztuk na palecie");
				header.createCell(20).setCellValue("angielska nazwa");
				header.createCell(21).setCellValue("hgw co to jest (skopiowane)");
				header.createCell(22).setCellValue("wieloskladnik");
				header.createCell(23).setCellValue("id parent");
				header.createCell(24).setCellValue("pozycja skladnika");
				
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


