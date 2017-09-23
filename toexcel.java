package traitement;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Locale;

import jxl.CellView;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.format.UnderlineStyle;
import jxl.write.Formula;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;


// Ce programme permet de prelever des donnees a partir d'un fichier.csv et d'ensuite extraire les donnees voulues dans un fichier excel.
// Il s'inspire et se refere au code de Lars Vogel (c) 2008, 2016 vogella GmbH Version 1.4, 29.08.2016
// Voici le lien qui mene a son code : http://www.vogella.com/tutorials/JavaExcel/article.html#installation ; Partie 2

public class toexcel {
	
	private WritableCellFormat timesBoldUnderline;
    private WritableCellFormat times;
    private String inputFile;
    
    public void setOutputFile(String inputFile) {
        this.inputFile = inputFile;
    }

    public ArrayList<String[]> extract(String nom) { // Fait par Younes Ijichi, s'inspire du code de mkyong
    												 // URL : https://www.mkyong.com/java/how-to-read-and-parse-csv-file-in-java/
		// INFORMATIONS METEO
		String[] jour = new String[59];
		String[] tempMoy = new String[59];
		String[] djc = new String[59];
		String[] neige = new String[59];
		ArrayList<String[]> donnees = new ArrayList<String[]>();

        String csvFile = "c:/temp/"+ nom +".csv";
        BufferedReader br = null;
        String line = "";
        ArrayList<String> lines = new ArrayList<String>();     
                  
        try {
            br = new BufferedReader(new FileReader(csvFile));
            while ((line = br.readLine()) != null) {          	
            	lines.add(line);
            }

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (br != null) {
                try {
                    br.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        
        // L'arraylist "lines" est remplie avec les donnees
        
        int index = 0;
        
        for (int i = 26; i < lines.size(); i++) {
        	String ligne = lines.get(i);
        	
        	String[] elements = ligne.split(",\"");
        	
        	if (elements[0].split("-")[1].equals("01") || elements[0].split("-")[1].equals("02")) {
        		jour[index] = elements[0].split("-")[2].split("\"")[0];
        		tempMoy[index] = elements[9].split("\"")[0];    
	        	djc[index] = elements[11].split("\"")[0];
	        	neige[index] = elements[21].split("\"")[0];

	        	index++;
        	}	
        }
        
        donnees.add(jour);
    	donnees.add(tempMoy);
    	donnees.add(djc);
    	donnees.add(neige);
    	
    	return donnees;
    }
    
    public void write() throws IOException, WriteException {
            File file = new File(inputFile);
            WorkbookSettings wbSettings = new WorkbookSettings();

            wbSettings.setLocale(new Locale("en", "EN"));

            WritableWorkbook workbook = Workbook.createWorkbook(file, wbSettings);
            workbook.createSheet("Statistiques", 0);
            WritableSheet excelSheet = workbook.getSheet(0);
            createLabel(excelSheet);
            createContent(excelSheet);

            workbook.write();
            workbook.close();
    }

    private void createLabel(WritableSheet sheet) // Fait par Lars Vogel (c) Vogella GmbH

                    throws WriteException {
            // Lets create a times font
            WritableFont times10pt = new WritableFont(WritableFont.TIMES, 10);
            // Define the cell format
            times = new WritableCellFormat(times10pt);
            // Lets automatically wrap the cells
            times.setWrap(true);

            // create create a bold font with underlines
            WritableFont times10ptBoldUnderline = new WritableFont(
                            WritableFont.TIMES, 10, WritableFont.BOLD, false,
                            UnderlineStyle.SINGLE);
            timesBoldUnderline = new WritableCellFormat(times10ptBoldUnderline);
            // Lets automatically wrap the cells
            timesBoldUnderline.setWrap(true);

            CellView cv = new CellView();
            cv.setFormat(times);
            cv.setFormat(timesBoldUnderline);
            cv.setAutosize(true);
    }

    private void createContent(WritableSheet sheet) throws WriteException, // Fait par Younes Ijichi, s'inspire de la methode de Lars Vogel (c) 2008, 2016 vogella GmbH
                    RowsExceededException {
    	
    		ArrayList<String[]> donnees = new ArrayList<String[]>();
    		
    		// Donnees 2010
    		
    		donnees = extract("2010");
    		
    		int index = 1;
    		
    		for(int k = 0; k < donnees.size(); k++){
    			String[] tab = donnees.get(k);
				for (int i = 0; i < tab.length; i++){
					if(i < 31) { // Janvier
						addLabel(sheet, index, i+2, tab[i]);
					}
					if(i >= 31) { // Fevrier
						addLabel(sheet, index, i+4, tab[i]);
					}
					
				}
				index++;
    		}
    		
    		donnees.clear();
    		
    		// Donnees 2017
    		
    		donnees = extract("2017");
    		
    		index = 6;
    		
    		for(int k = 0; k < donnees.size(); k++){
    			String[] tab = donnees.get(k);
				for (int i = 0; i < tab.length; i++){
					if(i < 31) { // Janvier
						addLabel(sheet, index, i+2, tab[i]);
					}
					if(i >= 31) { // Fevrier
						addLabel(sheet, index, i+4, tab[i]);
					}
					
				}
				index++;
    		}
    		
    		
    		addLabel(sheet, 1, 0, "Janvier 2010");
    		addLabel(sheet, 1, 33, "Fevrier 2010");
    		
    		addLabel(sheet, 6, 0, "Janvier 2017");
    		addLabel(sheet, 6, 33, "Fevrier 2017");
    	
    		addLabel(sheet, 1, 1, "Jour");
    		addLabel(sheet, 1, 34, "Jour");
    		addLabel(sheet, 6, 1, "Jour");
    		addLabel(sheet, 6, 34, "Jour");
    	
    		addLabel(sheet, 2, 1, "Temp. moy. °C");
    		addLabel(sheet, 2, 34, "Temp. moy. °C");
    		addLabel(sheet, 7, 1, "Temp. moy. °C");
    		addLabel(sheet, 7, 34, "Temp. moy. °C");
    		
    		addLabel(sheet, 3, 1, "DJC");
    		addLabel(sheet, 3, 34, "DJC");
    		addLabel(sheet, 8, 1, "DJC");
    		addLabel(sheet, 8, 34, "DJC");
    		
    		addLabel(sheet, 4, 1, "Neige au sol");
    		addLabel(sheet, 4, 34, "Neige au sol");
    		addLabel(sheet, 9, 1, "Neige au sol");
    		addLabel(sheet, 9, 34, "Neige au sol");
    		
    		addLabel(sheet, 0, 0, "No. de l'observation");
    		addLabel(sheet, 0, 33, "No. de l'observation");
    		
    		addLabel(sheet, 5, 0, "No. de l'observation");
    		addLabel(sheet, 5, 33, "No. de l'observation");
    		
    		for (int i = 0; i < 31; i++){ // Janvier 2010 et 2017
    			addNumber(sheet, 0, i+2, i+1);
    			addNumber(sheet, 5, i+2, i+1);
    		}    		
    		for (int i = 31; i < 59; i++){ // Fevrier 2010 et 2017
    			addNumber(sheet, 0, i+4, i+1);
    			addNumber(sheet, 5, i+4, i+1);
    		}
    		              
    }

    private void addCaption(WritableSheet sheet, int column, int row, String s) // Fait par Lars Vogel (c) Vogella GmbH
                    throws RowsExceededException, WriteException {
            Label label;
            label = new Label(column, row, s, timesBoldUnderline);
            sheet.addCell(label);
    }

    private void addNumber(WritableSheet sheet, int column, int row, // Fait par Lars Vogel (c) Vogella GmbH
                    Integer integer) throws WriteException, RowsExceededException {
            Number number;
            number = new Number(column, row, integer, times);
            sheet.addCell(number);
    }

    private void addLabel(WritableSheet sheet, int column, int row, String s) // Fait par Lars Vogel (c) Vogella GmbH
                    throws WriteException, RowsExceededException {
            Label label;
            label = new Label(column, row, s, times);
            sheet.addCell(label);
    }
    
    public static void main(String[] args) throws WriteException, IOException { // Fait par Lars Vogel (c) Vogella GmbH
        toexcel test = new toexcel();
        test.setOutputFile("c:/temp/lars.xls");
        test.write();
        System.out.println("Please check the result file under c:/temp/lars.xls ");
        
    }
    
}

