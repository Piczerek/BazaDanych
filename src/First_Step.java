// Piotr Misztal


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

public class First_Step {
int TableElement=0;  					//table iterator
int ExcelColumnQuantity;				//variable holding amount of excel columns 
int MonoComponentIndicatorColumn;		//variable holding number of column holding indicator if component is mono or poly
int ComponentColumn;					//variable holding number of column holding component
int RowToSaveIn=1;						//variable holding actual saving row number 
String WholeComponent;					//string holding component 
String ActualCharacter, 				//one character of "WholeComponent", actual character
ActualCharacterPlusOnePosition, 		//
ActualCharacterPlusTwoPositions,		//next characters os "WholeComponent"
ActualCharacterPlusThreePositions, 		//
ActualCharacterPlusFivePositions, 		//
Previous;								//
String CutComponent="";					// string containing ready to save text to excel
int ComponentId;						// largest component id 
String ProductName="";					//string holding actual components product name
int ProductNameColumn;					// variable holding number of column holding product name
int ProductIdColumn;					// variable holding number of column holding component id
double IdParent=-1;						// variable holding component's parent id
String CutComponentCleaningHelper="";	// string used to clean CutComponent befor saving
int Beggining=0;						//variable going through "WholeComponent" letters
boolean IfBeggining=true;				// boolean which tells us if now it is start of new poly component in mono component


//method which reads letter and checks if this letter is avaible if not returns empty string
 String ReadCharacter(int b){
	if(b<WholeComponent.length())
	{
		return Character.toString(WholeComponent.charAt(b));
	}
	return "";
}
	

//method used in comparision single "WholeComponent" letter with numbers
 static boolean IfEqualsNumber(String c){
	if (c.equals("0") || c.equals("1") || c.equals("2") || 
		c.equals("3") || c.equals("4") || c.equals("5") ||
		c.equals("6") || c.equals("7") || c.equals("8") || 
		c.equals("9"))
		{
			return true;
		}
	return false;
}
	

// method bases on table "tablica" values, depending on values creates specialised cell format and saves content of "tablica"
 static void zapisDoExcella(Row row, int i, struct tablica[]){
	if (tablica[i].text != "" && tablica[i].liczba == -1)
		row.createCell(i).setCellValue(tablica[i].text);
	else if (tablica[i].text == "" && tablica[i].liczba != -1)
		row.createCell(i).setCellValue(tablica[i].liczba);
	else
		row.createCell(i).setCellType(Cell.CELL_TYPE_BLANK);	
}
	
	
// method bases on excel cells content, reads a number cell in row number row and saves to table "tablica"	
 void czytaniezExcella(int a, Row row, struct tablica[]){
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
	
 int getTableElement() {
	return TableElement;
}


 void setTableElement(int tableElement) {
	TableElement = tableElement;
}


int getExcelColumnQuantity() {
	return ExcelColumnQuantity;
}


 void setExcelColumnQuantity(int excelColumnQuantity) {
	ExcelColumnQuantity = excelColumnQuantity;
}


 int getMonoComponentIndicatorColumn() {
	return MonoComponentIndicatorColumn;
}


 void setMonoComponentIndicatorColumn(int MonoComponentIndicatorColumn) {
	MonoComponentIndicatorColumn = MonoComponentIndicatorColumn;
}


int getComponentColumn() {
	return ComponentColumn;
}


 void setComponentColumn(int componentColumn) {
	ComponentColumn = componentColumn;
}


 int getRowToSaveIn() {
	return RowToSaveIn;
}


 void setRowToSaveIn(int rowToSaveIn) {
	RowToSaveIn = rowToSaveIn;
}


 String getWholeComponent() {
	return WholeComponent;
}


 void setWholeComponent(String wholeComponent) {
	WholeComponent = wholeComponent;
}


 String getActualCharacter() {
	return ActualCharacter;
}


 void setActualCharacter(String actualCharacter) {
	ActualCharacter = actualCharacter;
}


 String getActualCharacterPlusOnePosition() {
	return ActualCharacterPlusOnePosition;
}


 void setActualCharacterPlusOnePosition(String actualCharacterPlusOnePosition) {
	ActualCharacterPlusOnePosition = actualCharacterPlusOnePosition;
}


 String getActualCharacterPlusTwoPositions() {
	return ActualCharacterPlusTwoPositions;
}


 void setActualCharacterPlusTwoPositions(String actualCharacterPlusTwoPositions) {
	ActualCharacterPlusTwoPositions = actualCharacterPlusTwoPositions;
}


 String getActualCharacterPlusThreePositions() {
	return ActualCharacterPlusThreePositions;
}


 void setActualCharacterPlusThreePositions(String actualCharacterPlusThreePositions) {
	ActualCharacterPlusThreePositions = actualCharacterPlusThreePositions;
}


 String getActualCharacterPlusFivePositions() {
	return ActualCharacterPlusFivePositions;
}


 void setActualCharacterPlusFivePositions(String actualCharacterPlusFivePositions) {
	ActualCharacterPlusFivePositions = actualCharacterPlusFivePositions;
}


 String getPrevious() {
	return Previous;
}


 void setPrevious(String previous) {
	Previous = previous;
}


 String getCutComponent() {
	return CutComponent;
}


 void setCutComponent(String cutComponent) {
	CutComponent = cutComponent;
}


 int getComponentId() {
	return ComponentId;
}


 void setComponentId(int componentId) {
	ComponentId = componentId;
}


 String getProductName() {
	return ProductName;
}


 void setProductName(String productName) {
	ProductName = productName;
}


int getProductNameColumn() {
	return ProductNameColumn;
}


 void setProductNameColumn(int productNameColumn) {
	ProductNameColumn = productNameColumn;
}


 int getProductIdColumn() {
	return ProductIdColumn;
}


 void setProductIdColumn(int productIdColumn) {
	ProductIdColumn = productIdColumn;
}


 double getIdParent() {
	return IdParent;
}


 void setIdParent(double idParent) {
	IdParent = idParent;
}


 String getCutComponentCleaningHelper() {
	return CutComponentCleaningHelper;
}


 void setCutComponentCleaningHelper(String cutComponentCleaningHelper) {
	CutComponentCleaningHelper = cutComponentCleaningHelper;
}


 int getBeggining() {
	return Beggining;
}


 void setBeggining(int beggining) {
	Beggining = beggining;
}


 boolean isIfBeggining() {
	return IfBeggining;
}


 void setIfBeggining(boolean ifBeggining) {
	IfBeggining = ifBeggining;
}


//main
//public static void main(String[] args) {
public void StartScript(){
//scanner used to initialise our variables
	Scanner in = new Scanner(System.in);
	System.out.println("podaj ExcelColumnQuantity kolumn");
	ExcelColumnQuantity = in.nextInt();
	
	System.out.println("podaj w której kolumnie jest sklad do roztegowania");
	ComponentColumn = in.nextInt()-1;
	
	System.out.println("podaj w której kolumnie jest nazwa produktu");
	ProductNameColumn = in.nextInt()-1;
	
	System.out.println("podaj kolumne w ktorej jest wskaznik czy dany sklad jest MonoComponentIndicatorColumniem (0 jezeli nie ma takiej kolumny- wtedy rozbija kazdy sklad)");
	MonoComponentIndicatorColumn = in.nextInt()-1;
	
	System.out.println("podaj kolumne w ktorej jest id skladnika (jezeli nie ma 0)");
	ProductIdColumn = in.nextInt()-1;
	
	System.out.println("podaj najwieksze id produktu jaki znajduje sie aktualnie w bazie (lub jaki jest w danym excellu chodzi o unkalny identyfikator dla kazdego skladnika)");
	ComponentId = in.nextInt()+1;
//if in excel is no column with product id we create new column with our id's
	if(ProductIdColumn==-1)
	{
		ProductIdColumn=ExcelColumnQuantity+1;
	}
	 
	//initiation of table "tablica" which hac "ExcelColumnQuantity" length plus 7 for our script needs
	//tablica[ExcelColumnQuantity]- contents a place of poly component in mono component
	//tablica[ExcelColumnQuantity+1]- contents component id if excel didn t have it
	//tablica[ExcelColumnQuantity+2]- contents IdParent if needed
	//tablica[ExcelColumnQuantity+4]- position of character where component starts
	//tablica[ExcelColumnQuantity+5]- poly component length
	//initialisation of table "tablica" and "tablca_nazw" first will be changed in every loop, second is used once and it contents columns names
	struct[] tablica = new struct [ExcelColumnQuantity+7];
	struct[] tablica_nazw = new struct [ExcelColumnQuantity+6];

	//preparing table "tablica" to save in
	for(TableElement=0; TableElement<ExcelColumnQuantity+6; TableElement++)
	{ 
		tablica[TableElement]=new struct("",-1);
		tablica_nazw[TableElement]=new struct("",-1);
	}
	TableElement=0;
	
	try {		  
		// initialising inputstream
		InputStream is = new FileInputStream("D:/eclipse/BazaDanych-master/tescik.xlsx");
		//initialising excel workbook
		XSSFWorkbook workbook = new XSSFWorkbook(is);
		//initialising first sheet
        XSSFSheet sheet = workbook.getSheetAt(0);
        //creates new sheet "przed rozprocentowaniem"
        XSSFSheet sheet1 = workbook.createSheet("Przed rozprocentowaniem");
        //creates iterator for sheet zero
        java.util.Iterator<Row> rowIterator = sheet.iterator();
        System.out.println("First step started");
        
        Row row = rowIterator.next();
        //reads first row of excel sheet zero (columns names)
		for(TableElement=0; TableElement<ExcelColumnQuantity; TableElement++)
			{ 
				czytaniezExcella(TableElement,row, tablica_nazw);
			}
		TableElement=0;
		//additional names for which our script will generate content--------------------------------------------??????????add new columns
		tablica_nazw[ExcelColumnQuantity].text="pozycja skladnika w skladzie";
		tablica_nazw[ExcelColumnQuantity+1].text="id_skladnika";
		tablica_nazw[ExcelColumnQuantity+2].text="IdParent";

		
			
				
		while(rowIterator.hasNext()) {
			row = rowIterator.next();
			//clean all table "tablica" elements
			for(TableElement=0; TableElement<ExcelColumnQuantity; TableElement++)
			{ 
				tablica[TableElement].text="";
				tablica[TableElement].liczba=-1;
			}
			TableElement=0;
		
			//reads all cells in a row "row"			
			for(TableElement=0; TableElement<ExcelColumnQuantity; TableElement++)
				{ 
					czytaniezExcella(TableElement,row, tablica);
				}
			TableElement=0;
			//if there was no column with id product then assign to "tablica" new id
			if(ProductIdColumn==ExcelColumnQuantity+1)	
				{
					tablica[ProductIdColumn].liczba=ComponentId;
				}
			
			// assign content of excel cell which contains component to "WholeComponent"
			WholeComponent=tablica[ComponentColumn].text;
			//initialise WholeComponent length
			int dlugosc_slowa=WholeComponent.length();
			
			if(!ProductName.equals(tablica[ProductNameColumn].text))
			{
				ProductName=tablica[ProductNameColumn].text;
				tablica[ExcelColumnQuantity].liczba=1;
			}
			
			IdParent=tablica[ProductIdColumn].liczba;
			tablica[ExcelColumnQuantity+2].liczba=IdParent;
			//if ther is no column with mono/poly component mark do:
			if(MonoComponentIndicatorColumn==-1)
			{
				
			System.out.println("przygotuj wczesniej excella robiac odpowiednie oznaczenia (kolumna z oznaczeniem wielskladnik)");	
			// do nothing we have to prepare excel first in case of comments etc.
			}
			//if we got column with mono/poly mark then do:
			else
			{	
				
				//ok, we got marking column; 
				//our database squad decided that we need (before components cutting), whole mono component rewritten, so we go:
				//if component is marked as mono (1) rewrite it then do next steps
				if (tablica[MonoComponentIndicatorColumn].liczba==1)
			{
				Row header = sheet1.createRow(RowToSaveIn);
				RowToSaveIn++;
				
				for(TableElement=0; TableElement<ExcelColumnQuantity+1; TableElement++)
				{
					zapisDoExcella(header, TableElement, tablica);
				}
				// increase counter which tells us about position of component in whole squad
				tablica[ExcelColumnQuantity].liczba++;
			}
				//overwritting existing component id with mine. We can not save components with their parents id
				tablica[ProductIdColumn].liczba=ComponentId;
				ComponentId++;
				//again when we got mono component mark we have to cut it into poly components
				if (tablica[MonoComponentIndicatorColumn].liczba==1)
				{
					//we change mark to 0 because our components are poly
					tablica[MonoComponentIndicatorColumn].liczba=0;
					
					// main cutting algorithm we go throug string (whole component squad) letter after letter,
					//we are seraching for special signs like ",", ":", "(" etc. which are dividers between poly components
					//essential thing iss to catch mono components in mono components like we have 3bit
					// it has biscuit, so we got: 3bit: component a, component b, biscuit: biscuit component A, biscuit component B,
					// biscuit component C (for example can be lecithin): biscuit component C component A...
					//crazy, and we have to find them and mark as 1 and divide them properly
					for(int i=0;i<dlugosc_slowa; i++)
					{
						//assign i letter of WholeComponent to a variable
						ActualCharacter=ReadCharacter(i);
						//if it is last letter we don' t need whole algorithm. Add letter to poly component and save:
						if (i==dlugosc_slowa-1 && !CutComponent.equals(""))
						{
							//check if that last letter is not useless for us rubbish:
							if(!(ActualCharacter.equals(",") ||ActualCharacter.equals("(") ||
									ActualCharacter.equals(")") ||ActualCharacter.equals("[") ||
									ActualCharacter.equals("]") || ActualCharacter.equals(":") ||
									ActualCharacter.equals(";") ||ActualCharacter.equals("/")) )
							{
								//if not add letter to WholeComponent
								CutComponent=CutComponent+ActualCharacter;
							}
							//save WholeComponent to excel procedure:
							//create new row in which we can save
							Row header = sheet1.createRow(RowToSaveIn);
							RowToSaveIn++;
							//if first character in "CutComponent" is empty space skip it, if not do nothing
							if(Character.toString(CutComponent.charAt(0)).equals(" "))
							{
								for(int y=1; y<CutComponent.length();y++)
								{
									CutComponentCleaningHelper+=Character.toString(CutComponent.charAt(y));
								}
								CutComponent=CutComponentCleaningHelper;
								CutComponentCleaningHelper="";
							}
							//same like above in case of situation where component looks like:__component. We do not need empty spaces before
							if(Character.toString(CutComponent.charAt(0)).equals(" "))
							{
								for(int y=1; y<CutComponent.length();y++)
								{
									CutComponentCleaningHelper+=Character.toString(CutComponent.charAt(y));
								}
								CutComponent=CutComponentCleaningHelper;
								CutComponentCleaningHelper="";
							}
							tablica[ComponentColumn].text=CutComponent;
							CutComponent="";
							// now saving to excel function:
							for(TableElement=0; TableElement<ExcelColumnQuantity+1; TableElement++)
							{
								zapisDoExcella(header, TableElement, tablica);
							}
							//now we got to check before saving id parent if id parent is not actual component id:
							//example as above :..biscuit component C (for example can be lecithin): biscuit component C component A...
							//so biscuit component C with id xyz can not have in id parent id xyz coz component can not be parent for itself
							if(tablica[ProductIdColumn].liczba!=tablica[ExcelColumnQuantity+2].liczba)
							{
								zapisDoExcella(header, TableElement+1, tablica);
							}
							//after all we can save a position of first character of mono component
							zapisDoExcella(header, ExcelColumnQuantity+4, tablica);
							
							//seva length of poly component based on first and last character of poly component
							tablica[ExcelColumnQuantity+5].liczba=i-tablica[ExcelColumnQuantity+4].liczba+1;
							// variables actualistaion
							TableElement=0;
							tablica[ExcelColumnQuantity].liczba++;
							tablica[ExcelColumnQuantity+4].liczba=-1;
							//check if last character of poly component is " ". if yes, decrease length of this component by one
							if(Character.toString(tablica[ComponentColumn].text.charAt(tablica[ComponentColumn].text.length()-1)).equals(" "))
							{
								tablica[ExcelColumnQuantity+5].liczba--;
							}
							//save to excel length of poly component
							zapisDoExcella(header, ExcelColumnQuantity+5, tablica);
							//flag. tells us if poly component was save. after saving we have to find new poly component beggining
							IfBeggining=true;
							
							//checks if we increased id counter if not we do it and asign it to table tablica 
							if(tablica[ProductIdColumn].liczba==ComponentId)
							{
									ComponentId++;
							}
							tablica[ProductIdColumn].liczba=ComponentId;
							
						}
						//when we got squad sometimes we got bold etc components. it is bolded with: <...> 
						//it is not real component so we got to put it into rubbish:
						else if(ActualCharacter.equals("<"))
						{
							for(int iterator=i;iterator<WholeComponent.length();iterator++)
							{
								ActualCharacter = Character.toString(WholeComponent.charAt(iterator));
								if(ActualCharacter.equals(">"))
								{
									i=iterator;
									break;
								}
							}
						}
						//algorithm core:
						//so we check if character is not a separation sign
						//if yes we need to check some other parameters
						else if (ActualCharacter.equals(",") ||ActualCharacter.equals("(") ||
								ActualCharacter.equals(")") ||ActualCharacter.equals("[") ||
								ActualCharacter.equals("]") || ActualCharacter.equals(":") ||
								ActualCharacter.equals(";") ||ActualCharacter.equals("/"))
						{
							//we scan next letters. we don't move i iterator because we don't move in WholeComponent
							// we just need to check some conditions
							//letters avaibility is checked in "ReadCharacter" method
							ActualCharacterPlusOnePosition=ReadCharacter(i+1);
							ActualCharacterPlusTwoPositions=ReadCharacter(i+2);
							ActualCharacterPlusThreePositions=ReadCharacter(i+3);
							ActualCharacterPlusFivePositions=ReadCharacter(i+4);
							//now we check situation component xy,z% or component (xy,z%)
							//we don't want to separate component xy and z% ( component and xy and z%)
							//into two other rows so we have to check conditions if characters after actual are not numbers
							//we are using "IfEqualsNumber"
							//after this we got some conditions they are especially for one product: 5' di...smth...
							//this speciall component should be separated  despite it has number after comma
							if(!IfEqualsNumber(ActualCharacterPlusOnePosition) && !IfEqualsNumber(ActualCharacterPlusTwoPositions) && !CutComponent.equals("") && !CutComponent.equals(" ")||ActualCharacterPlusOnePosition.equals(" ")&& ActualCharacterPlusTwoPositions.equals("5")&& ActualCharacterPlusThreePositions.equals("'"))
							{	
								//clean WholeComponent
								Row header = sheet1.createRow(RowToSaveIn);
								RowToSaveIn++;
								if(Character.toString(CutComponent.charAt(0)).equals(" "))
								{
									for(int y=1; y<CutComponent.length();y++)
									{
										CutComponentCleaningHelper+=Character.toString(CutComponent.charAt(y));
									}
									CutComponent=CutComponentCleaningHelper;
									CutComponentCleaningHelper="";
								}
								if(Character.toString(CutComponent.charAt(0)).equals(" "))
								{
									for(int y=1; y<CutComponent.length();y++)
									{
										CutComponentCleaningHelper+=Character.toString(CutComponent.charAt(y));
									}
									CutComponent=CutComponentCleaningHelper;
									CutComponentCleaningHelper="";
								}
								
								//now we check if it is one of: "(","[" or ":"
								//it is important case in which we catch mono components in mono component (above lecithin example...)
								//if yes, we have to mark it as mono component ("1" in row where we got marks)
								//repeat it in excel (for data base needs) and then cut
								if(ActualCharacter.equals("(") || ActualCharacter.equals("[") || ActualCharacter.equals(":"))
								{
									CutComponentCleaningHelper=CutComponent;
									//we recognised mono component in mono component
									//we go through string and save it to CutComponentCleaningHelper string until we find bracket closing 
									//i found that ":" is used only with bracket or if those components are in this mono component till squad end
									for(int x=i; x<=WholeComponent.length();x++)
									{
										if( (ActualCharacter.equals("(") || ActualCharacter.equals("[")) && (!ReadCharacter(x).equals(")") &&
											!IfEqualsNumber(ReadCharacter(x-1)) || !(ReadCharacter(x).equals("]"))) || 
												ActualCharacter.equals(":"))
										CutComponentCleaningHelper+=ReadCharacter(x);
										else
											break;
									}
									
									//so we got whole monocomponent squad we save it:	
									tablica[ComponentColumn].text=CutComponentCleaningHelper;
									tablica[MonoComponentIndicatorColumn].liczba=1;
									for(TableElement=0; TableElement<ExcelColumnQuantity+1; TableElement++)
									{
										zapisDoExcella(header, TableElement, tablica);
									}
									
									if(tablica[ProductIdColumn].liczba!=tablica[ExcelColumnQuantity+2].liczba)
									{
										zapisDoExcella(header, TableElement+1, tablica);
									}
									// variables actualisation
									CutComponentCleaningHelper="";
									header = sheet1.createRow(RowToSaveIn);
									RowToSaveIn++;
									TableElement=0;
									tablica[ExcelColumnQuantity].liczba++;
									tablica[MonoComponentIndicatorColumn].liczba=0;
									ComponentId++;
									tablica[ProductIdColumn].liczba=ComponentId;
									
									
								}
								tablica[ComponentColumn].text=CutComponent;
								//here we save component. we checked that it is not comma before percentage mark, if it is monocomponent
								//we did rewritting whole monocomponent for database needs and now we write single component
								CutComponent="";
								for(TableElement=0; TableElement<ExcelColumnQuantity+1; TableElement++)
								{
									zapisDoExcella(header, TableElement, tablica);
								}
								
								if(tablica[ProductIdColumn].liczba!=tablica[ExcelColumnQuantity+2].liczba)
								{
									zapisDoExcella(header, TableElement+1, tablica);
								}
								zapisDoExcella(header, ExcelColumnQuantity+4, tablica);
								//variables actualisation
								tablica[ExcelColumnQuantity+5].liczba=i-tablica[ExcelColumnQuantity+4].liczba+1;
								tablica[ExcelColumnQuantity+4].liczba=-1;
								TableElement=0;
								tablica[ExcelColumnQuantity].liczba++;
								tablica[MonoComponentIndicatorColumn].liczba=0;
								//check if last character is not empty space if yes we decrese component length
								// it is prepared for situation like underlining some components in squad
								//we don't want to underline empty space at the end of component
								if(Character.toString(tablica[ComponentColumn].text.charAt(tablica[ComponentColumn].text.length()-1)).equals(" "))
								{
									tablica[ExcelColumnQuantity+5].liczba--;
								}
								//saving to excel and reset flag
								zapisDoExcella(header, ExcelColumnQuantity+5, tablica);
								IfBeggining=true;
								//if we divided squad on one of signs: "(","[" or":" we have to asign to it id parent
								//due to this asignment we know the tree which components created
								if(ActualCharacter.equals("(") || ActualCharacter.equals("[") || ActualCharacter.equals(":"))
								{
									tablica[ExcelColumnQuantity+2].liczba=tablica[ProductIdColumn].liczba;
									
								}
								//increase product id if was used
								if(tablica[ProductIdColumn].liczba==ComponentId)
								{
									ComponentId++;
								}
								//asign id product to next component
								tablica[ProductIdColumn].liczba=ComponentId;
								//when monocomponent in monocomponent ended we got to restore old id parent 
								if(ActualCharacter.equals(")") || ActualCharacter.equals("]") )
								{
									tablica[ExcelColumnQuantity+2].liczba=IdParent;
								}
							}
							//if we recognised numbers after comma (347 line) we got to add this letter and don't save
							else if (ActualCharacter.equals(",")&& !CutComponent.equals(" ")&& !CutComponent.equals(""))
							CutComponent=CutComponent+ActualCharacter;
							
						}
						//now if character was not special sign or sign with procent number
						// we can have situation "component A and component B" and we have to divide components 
						//connected with "and" (polish "i") or with " - "
						//it is below
						else if( ActualCharacter.equals(" ") )
						{
							ActualCharacterPlusOnePosition=ReadCharacter(i+1);
							ActualCharacterPlusTwoPositions=ReadCharacter(i+2);
							ActualCharacterPlusThreePositions=ReadCharacter(i+3);
							//here we check if it is space in many WholeComponent sigle component or if it is space before and/"-"
							if((ActualCharacterPlusOnePosition.equals("i") && ActualCharacterPlusTwoPositions.equals(" ") || ActualCharacterPlusOnePosition.equals("-") && ActualCharacterPlusTwoPositions.equals(" ")) && !(CutComponent.equals("") || CutComponent.equals(" "))&& !IfEqualsNumber(ActualCharacterPlusThreePositions))
							{
								//if yes we clear component save it and skip and/"-"
								Row header = sheet1.createRow(RowToSaveIn);
								RowToSaveIn++;
								if(Character.toString(CutComponent.charAt(0)).equals(" "))
								{
									for(int y=1; y<CutComponent.length();y++)
									{
										CutComponentCleaningHelper+=Character.toString(CutComponent.charAt(y));
									}
									CutComponent=CutComponentCleaningHelper;
									CutComponentCleaningHelper="";
								}
								if(Character.toString(CutComponent.charAt(0)).equals(" "))
								{
									for(int y=1; y<CutComponent.length();y++)
									{
										CutComponentCleaningHelper+=Character.toString(CutComponent.charAt(y));
									}
									CutComponent=CutComponentCleaningHelper;
									CutComponentCleaningHelper="";
								}
								
								tablica[ComponentColumn].text=CutComponent;
								CutComponent="";
								//saving
								for(TableElement=0; TableElement<ExcelColumnQuantity+1; TableElement++)
								{
									zapisDoExcella(header, TableElement, tablica);
								}
								
								if(tablica[ProductIdColumn].liczba!=tablica[ExcelColumnQuantity+2].liczba)
								{
									zapisDoExcella(header, TableElement+1, tablica);
								}
								zapisDoExcella(header, ExcelColumnQuantity+4, tablica);
								tablica[ExcelColumnQuantity+5].liczba=i-tablica[ExcelColumnQuantity+4].liczba+1;
								tablica[ExcelColumnQuantity+4].liczba=-1;
								TableElement=0;
								tablica[ExcelColumnQuantity].liczba++;
								tablica[MonoComponentIndicatorColumn].liczba=0;
								//checking length
								if(Character.toString(tablica[ComponentColumn].text.charAt(tablica[ComponentColumn].text.length()-1)).equals(" "))
								{
									tablica[ExcelColumnQuantity+5].liczba--;
								}
								zapisDoExcella(header, ExcelColumnQuantity+5, tablica);
								IfBeggining=true;
								
								if(tablica[ProductIdColumn].liczba==ComponentId)
								{
									ComponentId++;
								}
								
								tablica[ProductIdColumn].liczba=ComponentId;
								//skipping and
								i+=2;
							}
							//this is for situation when we had empty WholeComponent "CutComponent" we don t need to save empty strings so we just skip "and"
							//and add next letters
							else if((ActualCharacterPlusOnePosition.equals("i") && ActualCharacterPlusTwoPositions.equals(" ") || ActualCharacterPlusOnePosition.equals("-") && ActualCharacterPlusTwoPositions.equals(" ")) && (CutComponent.equals("") || CutComponent.equals(" ")))
									{
								i+=2;
									}
							//if in first condition we found that we had comma and number after comma there we add this actual character 
							//to our string "CutComponent
							else
								CutComponent=CutComponent+ActualCharacter;
						}
						//if it is space or comma and our WholeComponent is empty we need no spaces at the beggining or commas so we skip it
						else if ((ActualCharacter.equals(" ")|| ActualCharacter.equals(",")) && CutComponent.equals(""))
						{
									
						}
						//like above but about "*"
						else if (ActualCharacter.equals("*"))
						{
							
						}
						//if it is normal letter add it and! 
						//there are two important things first flag check:
						//if flag is false that means we got new poly component and we need it start letter position
						//second thing: if this letter was last we need to add it and save in some situations id did
						//not save single letter components like vitamin A, B, C that is for last single letter poly component
						else
						{
							CutComponent=CutComponent+ActualCharacter;
							if(IfBeggining==true)
							{
							Beggining=i;	
							tablica[ExcelColumnQuantity+4].liczba=Beggining;
							}
							IfBeggining=false;
							if (i==dlugosc_slowa-1 && !CutComponent.equals(""))
							{
								
								
								Row header = sheet1.createRow(RowToSaveIn);
								RowToSaveIn++;
								if(Character.toString(CutComponent.charAt(0)).equals(" "))
								{
									for(int y=1; y<CutComponent.length();y++)
									{
										CutComponentCleaningHelper+=Character.toString(CutComponent.charAt(y));
									}
									CutComponent=CutComponentCleaningHelper;
									CutComponentCleaningHelper="";
								}
								if(Character.toString(CutComponent.charAt(0)).equals(" "))
								{
									for(int y=1; y<CutComponent.length();y++)
									{
										CutComponentCleaningHelper+=Character.toString(CutComponent.charAt(y));
									}
									CutComponent=CutComponentCleaningHelper;
									CutComponentCleaningHelper="";
								}
								tablica[ComponentColumn].text=CutComponent;
								CutComponent="";
								
								for(TableElement=0; TableElement<ExcelColumnQuantity+1; TableElement++)
								{
									zapisDoExcella(header, TableElement, tablica);
								}
								
								if(tablica[ProductIdColumn].liczba!=tablica[ExcelColumnQuantity+2].liczba)
								{
									zapisDoExcella(header, TableElement+1, tablica);
								}
								zapisDoExcella(header, ExcelColumnQuantity+4, tablica);
								tablica[ExcelColumnQuantity+5].liczba=i-tablica[ExcelColumnQuantity+4].liczba+1;
								tablica[ExcelColumnQuantity+4].liczba=-1;
								TableElement=0;
								tablica[ExcelColumnQuantity].liczba++;
								
								if(Character.toString(tablica[ComponentColumn].text.charAt(tablica[ComponentColumn].text.length()-1)).equals(" "))
								{
									tablica[ExcelColumnQuantity+5].liczba--;
								}
								zapisDoExcella(header, ExcelColumnQuantity+5, tablica);
								IfBeggining=true;
								
								if(tablica[ProductIdColumn].liczba==ComponentId)
								{
										ComponentId++;
								}
								tablica[ProductIdColumn].liczba=ComponentId;
								
							}
						}		
					}	
				}
			//if it was polycomponent just save it
				else
				{
					Row header = sheet1.createRow(RowToSaveIn);
					RowToSaveIn++;
					for(TableElement=0; TableElement<ExcelColumnQuantity+1; TableElement++)
					{
						zapisDoExcella(header, TableElement, tablica);
					}
					tablica[ExcelColumnQuantity].liczba++;
				}
						
			}
							
	}
							

	//after all we use title table, save titles and finish program	
	Row header = sheet1.createRow(0);
	System.out.println("First step is saving its progress...");
	for(TableElement=0; TableElement<ExcelColumnQuantity+3; TableElement++)
		{
			zapisDoExcella(header, TableElement, tablica_nazw);
		}
	TableElement=0;
	//close all opened: sheet, workbook etc	
	//remove sheet from which we took data, we save them into second sheet
	int index = workbook.getSheetIndex(sheet);
	workbook.removeSheetAt(index);			
	is.close();
	FileOutputStream out = new FileOutputStream(new File("D:/eclipse/BazaDanych-master/tescik.xlsx"));
	workbook.write(out);
	out.close();
			
	} 
	//exception catches
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
