package com.testLab.excel;

import java.io.File;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ReadExcelFile {
	
	private String           sourceFile;
	private String           outputFile;
	private DataFormatter    dataFormatter;
	private SimpleDateFormat dateFormatter;
	private DecimalFormat    decimalFormatter;
	private String           pattern;
	
	public ReadExcelFile(String sourceFile) {
		
		this.sourceFile = sourceFile;
		
		pattern = "#0.####";
		decimalFormatter = new DecimalFormat(pattern);
		dataFormatter    = new DataFormatter();
		dateFormatter    = new SimpleDateFormat("dd/MM/yyyy");
		
		try {
			File excelFile = new File(sourceFile);
			Workbook workbook = WorkbookFactory.create(excelFile);
			
			for (Sheet sheet : workbook) {
				processSheet(sheet);
			}
			
		}catch(IOException ex1){
			System.out.println("Erreur de lecteur du fichier : " + ex1.getMessage());
		}catch(EncryptedDocumentException ex3){
			System.out.println("Le fichier Excel est crypte : " + ex3.getMessage());
		}
		
	}
	
	private void processSheet(Sheet sheet){
		
		System.out.println(sheet.getSheetName());
		
		for (int i = sheet.getFirstRowNum(); i <= sheet.getLastRowNum(); i++) {
			Row row = sheet.getRow(i);
			for (int j = row.getFirstCellNum(); j <= row.getLastCellNum(); j++) {
				Cell cell = row.getCell(j);
				String 	cellValue = dataFormatter.formatCellValue(cell);
				String result     = getCellValue(cell);
				//System.out.print(String.format("%-20.20s", cellValue) + " | ");
				System.out.print(String.format("%-20.20s", result) + " | ");
			}
			System.out.print("\n");
		}
		System.out.print("\n");
	}
	
	private  String getCellValue(Cell cell){
		if (cell != null) {
		
			CellType cellType = cell.getCellType();
			
			if (cellType.equals(CellType.FORMULA)) {
				cellType = cell.getCachedFormulaResultType();
			}
			
			switch (cellType){
				case BOOLEAN:
					return String.valueOf( cell.getBooleanCellValue() );
				case NUMERIC:
					if (DateUtil.isCellDateFormatted(cell)){
						return String.valueOf(dateFormatter.format(cell.getDateCellValue()));
					}else{
						return decimalFormatter.format(cell.getNumericCellValue());
					}
				case STRING:
					return cell.getRichStringCellValue().toString();
				default:
					break;
			}
		}
		return "   ";
	}
		
}
