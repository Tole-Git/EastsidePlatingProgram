package excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.Scanner;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class MyExcel {
	
	private static Row row;
	
	public static void main(String[] args) {
		
		mainMenu();
		
	}
	
	public static void mainMenu() {
		 Scanner userInput = new Scanner(System.in);
			
		System.out.println("*Main Menu* Please enter an option.");
		System.out.println("1 : viewData");
		System.out.println("2 : newData");
		System.out.println("3 : exit");
		System.out.print("user : ");
		double fnum = userInput.nextDouble();
		
		if(fnum == 1) {
			readExcel();
		} 
		if(fnum == 2) {
			verifyExcel(userInput);
		} 
		if(fnum == 3) {
			System.out.println();
			System.out.println("Goodbye :-)");
		} 
	}
	
	public static void verifyExcel(Scanner console) {
		System.out.print("PartNo1 : ");
		String part = console.next();
		int checked = checkExcel(part);
		
		if (checked == 1000) {
			writeExcel(part);
		} else {
			System.out.println("Existing part.");
			writeExcel(checked);
			}
		
	}
	
	public static int checkExcel(String partN1) {
		int rowNumber = 1000;
		
		try {
			FileInputStream file = new FileInputStream(new File("testdata1.xlsx"));

			XSSFWorkbook workbook = new XSSFWorkbook(file);
			XSSFSheet sheet = workbook.getSheetAt(0);
			Cell cell = null;
			int noOfRows = sheet.getLastRowNum();
			Scanner userInput = new Scanner(System.in);

		
			for(int i = 0; i <= noOfRows; i++) {
				row = sheet.getRow(i);
				cell= row.getCell(0);
				String compare = cell.toString();
		
				if(compare.equals(partN1) == true) {
					rowNumber = i;
					return rowNumber;
				}
			}
		
		file.close();
		
		FileOutputStream outFile =new FileOutputStream(new File("testdata1.xlsx"));
		workbook.write(outFile);
		outFile.close();
		
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
		e.printStackTrace();
		}
		return rowNumber;
		
	}
	
	public static void writeExcel(String partNo1) {
		try {
			FileInputStream file = new FileInputStream(new File("testdata1.xlsx"));

			XSSFWorkbook workbook = new XSSFWorkbook(file);
			XSSFSheet sheet = workbook.getSheetAt(0);
			Cell cell = null;
			int noOfRows = sheet.getLastRowNum();
			int noOfCol = 4;
			int totalQuantity = 0;
			int boxQuantity = 0;
			double boxWeight = 0;
			String weightPerPart;
			Scanner userInput = new Scanner(System.in);
			
			row = sheet.createRow(noOfRows + 1);		
			cell= row.createCell(0);
			cell.setCellValue(partNo1);
			
			//Update the value of cell
			for(int i = 1; i<= noOfCol; i++) {				

				if(i == 1) {
					System.out.print("PartNo2 : ");
					cell= row.createCell(i);
					cell.setCellValue(userInput.nextLine());
				}
				if(i == 2) {
					System.out.print("boxes : ");
					cell= row.createCell(i);
					cell.setCellValue(userInput.nextLine());
				}
				if(i == 3) {
					System.out.print("Box Quantity : ");
					boxQuantity = userInput.nextInt();
					System.out.print("Box Weight : ");
					boxWeight = userInput.nextDouble();
					System.out.print("Weight per Box : " );
					weightPerPart = String.format("%.4f", boxWeight / boxQuantity);
					System.out.println(weightPerPart);
					cell= row.createCell(5);
					cell.setCellValue(weightPerPart);
					System.out.print("Total Quantity : ");
					cell= row.createCell(i);
					totalQuantity = userInput.nextInt();
					cell.setCellValue(totalQuantity);
					cell= row.createCell(4);
					cell.setCellValue(totalQuantity * Double.parseDouble(row.getCell(5).toString()));

					System.out.print("Total Weight : ");
					System.out.println(row.getCell(4).toString());
					System.out.print("Pallet Number : ");
					cell = row.createCell(6);
					cell.setCellValue(userInput.nextInt());
					System.out.println("Press enter to continue...");
			        try {
			            int read = System.in.read(new byte[2]);
			        } catch (IOException e) {
			            e.printStackTrace();
			        }
				}
				if(i == 4) {
					System.out.println("Row Completed.");	
				}
				
			}
			file.close();
			
			FileOutputStream outFile =new FileOutputStream(new File("testdata1.xlsx"));
			workbook.write(outFile);
			outFile.close();
			
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		System.out.println();
		mainMenu();
	}
	
	public static void writeExcel(int existing) {
		try {
			FileInputStream file = new FileInputStream(new File("testdata1.xlsx"));

			XSSFWorkbook workbook = new XSSFWorkbook(file);
			XSSFSheet sheet = workbook.getSheetAt(0);
			Cell cell = null;
			int noOfRows = sheet.getLastRowNum();
			int noOfCol = 4;
			double totalQuantity = 0;
			int boxQuantity = 0;
			double boxWeight = 0;
			double weightPerPart;
			Scanner userInput = new Scanner(System.in);
			
			row = sheet.getRow(existing);		
			
			//Update the value of cell
			for(int i = 2; i<= noOfCol; i++) {				

				if(i == 2) {
					System.out.println("Pallet No. " + row.getCell(6).toString());
					System.out.print("Add boxes : ");
					double box = Double.parseDouble(row.getCell(2).toString()) + userInput.nextInt();
					cell= row.createCell(i);
					cell.setCellValue(box);				}
				if(i == 3) {
					System.out.print("Add Quantity : ");
					totalQuantity = Double.parseDouble(row.getCell(3).toString()) + userInput.nextInt();
					
					cell= row.createCell(3);
					cell.setCellValue(totalQuantity);
					
					cell= row.createCell(4);
					cell.setCellValue(String.format("%.2f", totalQuantity * Double.parseDouble(row.getCell(5).toString())));
					
					System.out.print("Total Weight : ");
					System.out.println(row.getCell(4).toString());
					System.out.println("Press enter to continue...");
			        try {
			            int read = System.in.read(new byte[2]);
			        } catch (IOException e) {
			            e.printStackTrace();
			        }
				}
				if(i == 4) {
					System.out.println("Row Completed.");	
				}
				
			}
			file.close();
			
			FileOutputStream outFile =new FileOutputStream(new File("testdata1.xlsx"));
			workbook.write(outFile);
			outFile.close();
			
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		System.out.println();
		mainMenu();
	}
	
	public static void readExcel() {
	try {
		
		FileInputStream file = new FileInputStream(new File("testdata1.xlsx"));
		
		//Get the workbook instance for XLS file 
		XSSFWorkbook workbook = new XSSFWorkbook(file);

		//Get first sheet from the workbook
		XSSFSheet sheet = workbook.getSheetAt(0);
		
		//Iterate through each rows from first sheet
		Iterator<Row> rowIterator = sheet.iterator();
		while(rowIterator.hasNext()) {
			Row row = rowIterator.next();
			
			//For each row, iterate through each columns
			Iterator<Cell> cellIterator = row.cellIterator();
			while(cellIterator.hasNext()) {
				
				Cell cell = cellIterator.next();
				
				switch(cell.getCellType()) {
					case BOOLEAN:
						System.out.print(cell.getBooleanCellValue() + "\t\t");
						break;
					case NUMERIC:
						System.out.print(cell.getNumericCellValue() + "\t\t");
						break;
					case STRING:
						System.out.print(cell.getStringCellValue() + "\t\t");
						break;
				}
			}
			System.out.println("");
		}
		file.close();
		FileOutputStream out = 
			new FileOutputStream(new File("testdata1.xlsx"));
		workbook.write(out);
		out.close();
		
	} catch (FileNotFoundException e) {
		e.printStackTrace();
	} catch (IOException e) {
		e.printStackTrace();
	}
	System.out.println();
	mainMenu();
	}

}
