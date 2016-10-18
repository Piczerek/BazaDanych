
public class main {

	public static void main(String[] args) throws Exception {
		czytanie czyt = new czytanie();
		czyt.main(args);
		czyt=null;
		Stringchange rozprocentowanie = new Stringchange();
		rozprocentowanie.main(args);
		rozprocentowanie=null;
		System.out.println("GUNWO JESZCZE");
		czyszczenie czysc = new czyszczenie();
		czysc.main(args);
		czysc=null;
		}

	}
	
