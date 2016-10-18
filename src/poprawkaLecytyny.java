import java.io.*;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class poprawkaLecytyny {
		static double product_item_id	=0;
		static double product_item_type=0;
		static double product_item_position=0;
		static double product_item_id_parent=0;
		static double product_item_id_product =0;
		
		static String product_item_name="";
		static String product_item_quantity="";
		static String product_name="";
	
	
		static double product_item_id1	=0;
		static double product_item_type1=0;
		static double product_item_position1=0;
		static double product_item_id_parent1=0;
		static double product_item_id_product1 =0;
		
		static String product_item_name1="";
		static String product_item_quantity1="";
		static String product_name1="";
		
		static int i=0;
		static int a=0;
		static double position=1;
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
					DataFormatter formatter = new DataFormatter(); //creating formatter using the default locale
					 String j_username = formatter.formatCellValue(cell); 
					 return j_username;
					//double wartosc01=cell.getNumericCellValue();
					//long x = (long) wartosc01;
					//return Long.toString(x);
			
				case Cell.CELL_TYPE_STRING:
					return cell.getStringCellValue();
					
				case Cell.CELL_TYPE_FORMULA:
					return "";
					//return cell.getStringCellValue();
			
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
				i++;

				product_item_id					=	czytanieZExcellaIntegera(0, row);
				product_item_name				=	czytanieZExcellaStringa(1, row);
				product_item_type				=	czytanieZExcellaIntegera(2, row);
				product_item_position			=	czytanieZExcellaIntegera(3, row);
				product_item_id_parent			=	czytanieZExcellaIntegera(4, row);
				product_item_quantity			=	czytanieZExcellaStringa(5, row);
				product_item_id_product			=	czytanieZExcellaIntegera(6, row);
				product_name					=	czytanieZExcellaStringa(7, row);

				
				
				if(rowIterator.hasNext())
				{
					Row row1 = sheet.getRow(i);
				
				//product_item_id1				=	czytanieZExcellaIntegera(0, row1);
				product_item_name1				=	czytanieZExcellaStringa(1, row1);
				//product_item_type1				=	czytanieZExcellaIntegera(2, row1);
				//product_item_position1			=	czytanieZExcellaIntegera(3, row1);
				//product_item_id_parent1			=	czytanieZExcellaIntegera(4, row1);
				//product_item_quantity1			=	czytanieZExcellaStringa(5, row1);
				//product_item_id_product1		=	czytanieZExcellaIntegera(6, row1);
				//product_name1					=	czytanieZExcellaStringa(7, row1);
				}
				if(product_item_name.equals("lecytyny")||product_item_name.equals("lecytyny "))
				{
					if(product_item_name1.equals("z soi")|| product_item_name1.equals("z SOI") || product_item_name1.equals("soja")|| product_item_name1.equals("SOJA"))
					{
						product_item_name="lecytyny z soi";
						product_item_type=0;	
					}
				}
				
				if(product_item_position<position)
				{
				position=1;
						
				}
				
				if(!(product_item_name.equals("z soi")|| product_item_name.equals("z SOI") || product_item_name.equals("soja")))
				{
					Row header = sheet1.createRow(a);

					zapisDoExcella(0,	 "",					product_item_id,			 header);
					zapisDoExcella(1,	 product_item_name,		-1,	 						 header);
					zapisDoExcella(2,	 "",					product_item_type,			 header);
					zapisDoExcella(3,	 "",					position,					 header);
					zapisDoExcella(4,	 "",					product_item_id_parent,	 	 header);
					zapisDoExcella(5,	 product_item_quantity,	-1,	 						 header);
					zapisDoExcella(6,	 "",					product_item_id_product,	 header);//////
					zapisDoExcella(7,	 product_name,			-1,							 header);
				position++;
				a++;
				System.out.println(a);
				}
			}
//koniec programu	
				Row header = sheet1.createRow(0);
				header.createCell(0).setCellValue("product_item_id");
				header.createCell(1).setCellValue("product_item_name");
				header.createCell(2).setCellValue("product_item_type");
				header.createCell(3).setCellValue("position");
				header.createCell(4).setCellValue("product_item_id_parent");
				header.createCell(5).setCellValue(" product_item_quantity");
				header.createCell(6).setCellValue("product_item_id_product");
				header.createCell(7).setCellValue(" product_name");
				
				
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




