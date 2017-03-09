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

public class oznaczanie {
static int oczko_tablicy=0;
static int ilosc=10;
static int wieloskladnik=-1;
static int roztegowywany=0;
static int Row_to_save_in=1;
static boolean czy_wieloskladnik = false;
static String slowo;
static String aktualny_1, aktualny_2, aktualny_3, aktualny_4, aktualny_5, poprzedni;
static String pociete="";
static int numer_skladnika=75289;
static int pozycja_w_skladzie=1;
static String nazwa_produktu="";
static int nr_nazwa_produktu=-1;
static int kolumna_z_id_produktow=-1;
static double ID_PARENT=-1;
static boolean czy_uzyc_parent=true;
static String pomocnicza="";
static int poczatek=0;
static boolean czy_poczatek=true;



private static String czytanieliter(int b){
	if(b<slowo.length())
	{
		return Character.toString(slowo.charAt(b));
	}
	return "";
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
	


private static void zapisDoExcella(Row row, int i, struct tablica[]){
	if (tablica[i].text != "" && tablica[i].liczba == -1)
		row.createCell(i).setCellValue(tablica[i].text);
	else if (tablica[i].text == "" && tablica[i].liczba != -1)
		row.createCell(i).setCellValue(tablica[i].liczba);
	else
		row.createCell(i).setCellType(Cell.CELL_TYPE_BLANK);	
}
	
	
	
private static void czytaniezExcella(int a, Row row, struct tablica[]){
	Cell cell = row.getCell(a);
 if (cell!=null)
	{
		switch(cell.getCellType()) {		
			case Cell.CELL_TYPE_NUMERIC:
				System.out.println(cell.getNumericCellValue());
				tablica[oczko_tablicy].liczba= cell.getNumericCellValue();
				break;
			case Cell.CELL_TYPE_STRING:	
				tablica[oczko_tablicy].text=cell.getStringCellValue();
				break;
		}
	}	
}
	

public static void main(String[] args) {
		
	Scanner in = new Scanner(System.in);
	System.out.println("podaj ilosc kolumn");
	ilosc = in.nextInt();
	
	System.out.println("podaj w której kolumnie jest sklad do roztegowania");
	roztegowywany = in.nextInt()-1;
	
	System.out.println("podaj w której kolumnie jest nazwa produktu");
	nr_nazwa_produktu = in.nextInt()-1;
	
	System.out.println("podaj kolumne w ktorej jest wskaznik czy dany sklad jest wieloskladnikiem (0 jezeli nie ma takiej kolumny- wtedy rozbija kazdy sklad)");
	wieloskladnik = in.nextInt()-1;
	
	System.out.println("podaj kolumne w ktorej jest id skladnika (jezeli nie ma 0)");
	kolumna_z_id_produktow = in.nextInt()-1;
	
	System.out.println("podaj najwieksze id produktu jaki znajduje sie aktualnie w bazie (lub jaki jest w danym excellu chodzi o unkalny identyfikator dla kazdego skladnika)");
	numer_skladnika = in.nextInt()+1;
	
	if(kolumna_z_id_produktow==-1)
	{
		kolumna_z_id_produktow=ilosc+1;
	}
	 
	//tablica[ilosc]- pozycja skladnika w danym produkcie
	//tablica[ilosc+1]- pozycja skladnika w bazie danych jezeli nie ma wczesniej takiej kolumny
	//tablica[ilosc+2]- id_parent
	//tablica[ilosc+4]- czy wieloskladnik czy nie

	struct[] tablica = new struct [ilosc+6];
	struct[] tablica_nazw = new struct [ilosc+6];
	System.out.println("1");

		
	for(oczko_tablicy=0; oczko_tablicy<ilosc+5; oczko_tablicy++)
	{ 
		System.out.println(oczko_tablicy);
		tablica[oczko_tablicy]=new struct("",-1);
		tablica_nazw[oczko_tablicy]=new struct("",-1);
	}
	oczko_tablicy=0;
	
	try {		  
// tworzy input stream 
		InputStream is = new FileInputStream("D:\\tescik.xlsx");
//sczytuje excela
		XSSFWorkbook workbook = new XSSFWorkbook(is);
//bierze pierwszy arkusz
        XSSFSheet sheet = workbook.getSheetAt(0);
//tworzy nowy arkusz o nazwie "przed rozprocentowaniem
        XSSFSheet sheet1 = workbook.createSheet("Przed rozprocentowaniem");
//tworzy iterator dla arkusza zerowego
        java.util.Iterator<Row> rowIterator = sheet.iterator();
        
        
        Row row = rowIterator.next();			
		for(oczko_tablicy=0; oczko_tablicy<ilosc; oczko_tablicy++)
			{ 
				czytaniezExcella(oczko_tablicy,row, tablica_nazw);
			}
		oczko_tablicy=0;
		tablica_nazw[ilosc].text="pozycja skladnika w skladzie";
		tablica_nazw[ilosc+1].text="id_skladnika";
		tablica_nazw[ilosc+2].text="id_parent";

		
			
				
		while(rowIterator.hasNext()) {
			row = rowIterator.next();
//czyœci wszystkie oczka tablicy po kolei wierszy
			for(oczko_tablicy=0; oczko_tablicy<ilosc; oczko_tablicy++)
			{ 
				tablica[oczko_tablicy].text="";
				tablica[oczko_tablicy].liczba=-1;
			}
		oczko_tablicy=0;
		
// czyta po kolei kolumny w danym rzedzie			
			for(oczko_tablicy=0; oczko_tablicy<ilosc; oczko_tablicy++)
				{ 
					czytaniezExcella(oczko_tablicy,row, tablica);
				}
			oczko_tablicy=0;
			
				if(kolumna_z_id_produktow==ilosc+1)	
				{
					tablica[kolumna_z_id_produktow].liczba=numer_skladnika;
				}
			
			
			slowo=tablica[roztegowywany].text;
			int dlugosc_slowa=slowo.length();
			if(tablica[wieloskladnik].liczba!=0 && tablica[wieloskladnik].liczba!=2)
			{
			for(int i=0; i<dlugosc_slowa;i++)
			{
				aktualny_1=czytanieliter(i);
				if(aktualny_1.equals(",") && !IfEqualsNumber(czytanieliter(i+1))||aktualny_1.equals("(")&& !IfEqualsNumber(czytanieliter(i+1)) ||
								aktualny_1.equals(")") ||aktualny_1.equals("[")&& !IfEqualsNumber(czytanieliter(i+1)) ||
								aktualny_1.equals("]") || aktualny_1.equals(":") ||
								aktualny_1.equals(";") ||aktualny_1.equals("/")||aktualny_1.equals("-") && czytanieliter(i+1).equals(" ") && czytanieliter(i-1).equals(" ")
								||aktualny_1.equals("i") && czytanieliter(i+1).equals(" ") && i!=0 && czytanieliter(i-1).equals(" "))
				{
					tablica[wieloskladnik].liczba=1;
				
					
				}
			}
			if (!(tablica[wieloskladnik].liczba==1))
			{
				for(int i=0; i<dlugosc_slowa;i++)
				{
					aktualny_1=czytanieliter(i);
					if(aktualny_1.equals("%"))
					{
						tablica[wieloskladnik].liczba=1;
					
						
					}
			}
			}
			}
			Row header = sheet1.createRow(Row_to_save_in);
			Row_to_save_in++;
			for(oczko_tablicy=0; oczko_tablicy<ilosc+1; oczko_tablicy++)
			{
				zapisDoExcella(header, oczko_tablicy, tablica);
			}
			
			
		
			
			
			
			
			
			
			
			
	}
							

//koniec programu	
	Row header = sheet1.createRow(0);
	for(oczko_tablicy=0; oczko_tablicy<ilosc+3; oczko_tablicy++)
		{
			zapisDoExcella(header, oczko_tablicy, tablica_nazw);
			System.out.println("zapisuje");
		}
	oczko_tablicy=0;
			
	int index = workbook.getSheetIndex(sheet);
	workbook.removeSheetAt(index);			
	is.close();
	FileOutputStream out = new FileOutputStream(new File("D:\\tescik.xlsx"));
	workbook.write(out);
	out.close();
			
	} 
   	catch (FileNotFoundException e) 
		{
   			e.printStackTrace();
   			System.out.println("wyjebawszy1");
		} 
   	catch (IOException e) 
   		{
   			e.printStackTrace();
   			System.out.println("wyjebawszy2");
   		}
	
	
}
}
