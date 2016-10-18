import java.io.*;

import javax.swing.text.html.HTMLDocument.Iterator;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.usermodel.*;

public class Stringchange
{
	static String slowo=null;
	static String pociete="";
	static String c = null;
	static String d=null;
	static String a=null;
	static String e=null;
	
	static double Wieloskladnik;
	static double Product_item;
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
				Id_Parent				=	czytanieZExcellaIntegera(11, row);
				Pelen_sklad				=	czytanieZExcellaStringa	(13, row);
			
				slowo=Skladnik;
				
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
							pociete=Skladnik;
						}
						
						Row header = sheet1.createRow(rzad);
						zapisDoExcella(0,	 "",			 	Wieloskladnik,			 header);						
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
						zapisDoExcella(12,	 procent, 			-1,						 header);
						zapisDoExcella(13,	 Pelen_sklad,		-1,						 header);

						rzad++;
						procent="";
						pociete="";
			}
			
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
			header.createCell(12).setCellValue("procent");
			header.createCell(13).setCellValue("pelen sklad wieloskladnika");
			
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



