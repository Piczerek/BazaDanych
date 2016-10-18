import java.io.*;

import javax.swing.text.html.HTMLDocument.Iterator;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.usermodel.*;

public class stringchangeDaryNatury
{
	static String slowo=null;
	static String pociete="";
	static String c = null;
	static String d=null;
	static String a=null;
	static String e=null;
	
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
	static String pelen_sklad;
	


	
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
	static double Wieloskladnik;
	static double Id_Parent;
	static double Pozycja_skladnika;
	static String procent="";
	static String f="";
	
	static int i=0;
	static int zmiennaprocent=0;
	static int zmiennaspacja=0;
	static int zmienna1=0;
	static int pomocniczadoparent=0;
	static int rzad=0;
	
	static String czytanieZExcellaStringa(int a, Row row){
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
	
	static double czytanieZExcellaIntegera(int a, Row row){
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
	
	static void zapisDoExcella(int a, String wartosc, double wartosc01, Row row){
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
			//XSSFSheet sheet0 = workbook.getSheetAt(0);
			XSSFSheet sheet	 = workbook.getSheetAt(0);
			XSSFSheet sheet1 = workbook.createSheet("rozprocentowane");
			//Create a new row in current sheet
	
			//Iterate through each rows from first sheet
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
				Wieloskladnik			=	czytanieZExcellaIntegera(22, row);
				Id_Parent				=	czytanieZExcellaIntegera(23, row);
				Pozycja_skladnika		=	czytanieZExcellaIntegera(24, row);
				pelen_sklad				=	czytanieZExcellaStringa(25, row);
				
				slowo=skladniki;
				
// program idzie po kolejnych znakach Skladnik'a. gdy natrafi na procent zapisuje miejsce procenta
//cofa sie az natrafi na spacje lub na poczatek Skladnik'a. w zaleznosci czy procent jest 
// na poczatku czy na koncu skladnika'a odpowiednio wycina procent a nastepnie wstawia do excella
						if (Wieloskladnik!=3)
						{
						for (i=0; i<=slowo.length(); i++)
						{	
							if(i+1<=slowo.length())
							{
								a= Character.toString(slowo.charAt(i));
							}
							if (a.equals("%"))
							{
								zmienna1=1;
								zmiennaprocent=i;
								break;
							}
							else
							{
								zmienna1=0;
								procent="";
							}
						
						}
						
						if (zmienna1==1)
						{
							for(i=zmiennaprocent; i>0;i--)
							{
								if(i+1<slowo.length())
								{
									a= Character.toString(slowo.charAt(i));
									d=Character.toString(slowo.charAt(i+1));
								}
								if (a.equals(" ") && !(d.equals("%")))
								{
									zmiennaspacja=i;
									break;
								}
								else if(i==1)
								{
									zmiennaspacja=0;
								}
								
							}
						
							if (zmiennaspacja!=0)
							{
								for(i=0;i<zmiennaspacja;i++)
								{
									if(i+1<=slowo.length())
									{
										a= Character.toString(slowo.charAt(i));
									}
								pociete=pociete+a;
								}
							
								for(i=zmiennaspacja+1;i<=zmiennaprocent;i++)
								{
									if(i+1<=slowo.length())
									{
										a= Character.toString(slowo.charAt(i));
									}
								procent=procent+a;
								
								}
							}
							else
							{
								for(i=0;i<zmiennaprocent; i++)
								{
									if(i+1<=slowo.length())
									{
										a= Character.toString(slowo.charAt(i));
									}
								procent=procent+a;
								}
							if(zmiennaprocent+1<slowo.length())
								{
									for(i=zmiennaprocent+1; i<=slowo.length();i++)
										{
											if(i+1<=slowo.length())
												{
													a= Character.toString(slowo.charAt(i));
													pociete=pociete+a;
												}
										
										}
								}
							}
						}
						else
							{
								pociete=slowo;
								procent="";
							}
						
						String pomocniczy="";
						for (char a: procent.toCharArray())
							{	
								String b=Character.toString(a);
								if (!( b.equals("(") || b.equals(")")) )
									pomocniczy=pomocniczy+a;
							
							}
						procent=pomocniczy;
						pomocniczy="";
						}
						else 
						{
							pociete=skladniki;
						}
						
						Row header = sheet1.createRow(rzad);
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
						zapisDoExcella(23,	 "",		Id_Parent,						 header);
						zapisDoExcella(24,	 "",		Pozycja_skladnika,						 header);
						zapisDoExcella(25,	 procent, 			-1,						 header);


						rzad++;
						procent="";
						pociete="";
			}
			
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
			header.createCell(25).setCellValue("zwartosc procentowa");
			
			
			int index = workbook.getSheetIndex(sheet);
			workbook.removeSheetAt(index);
			file.close();
			FileOutputStream out = new FileOutputStream(new File("D:\\test.xlsx"));
			workbook.write(out);
			out.close();
			System.out.println("ZROBIONE, nie otwieraj jeszcze bazy");
			
		} 
	   	catch (FileNotFoundException e) {
			e.printStackTrace();
		} 
	   catch (IOException e) {
			e.printStackTrace();
		}
   }



}



