
public class main {

	public static void main(String[] args) throws Exception {

		First_Step first= new First_Step();
		first.StartScript();
		
		Second_Step second = new Second_Step();
		second.porzadkowanie(first.getExcelColumnQuantity(), first.getComponentColumn(), first.getMonoComponentIndicatorColumn());
		second=null;	
		
		Third_Step third = new Third_Step();
		third.porzadkowanie(first.getExcelColumnQuantity(), first.getComponentColumn(), first.getMonoComponentIndicatorColumn());
		third=null;
		
		System.out.println("The end");
		
		}

	}
	
