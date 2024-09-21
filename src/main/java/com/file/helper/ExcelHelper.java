package com.file.helper;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;

import com.file.model.FinancialRecords;

public class ExcelHelper {
	public static String XLSX_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
	public static String XLS_TYPE = "application/vnd.ms-excel"; // for .xls files
	public static String[] SUPPORTED_TYPES = {XLSX_TYPE, XLS_TYPE};
	  static String SHEET = "Financials";


	public static boolean hasExcelFormat(MultipartFile file) {
	    String fileType = file.getContentType();

	    System.out.println("File Content Type: " + file.getContentType());

		for (String type : SUPPORTED_TYPES) {
	        if (type.equals(fileType)) {
	            return true;
	        }
	    }
	    
	    return false;
	}
	
	public static ByteArrayInputStream finanacialRecorsdToExcel(List<FinancialRecords>financialRecordsList) throws IOException {
		Workbook workbook = null;
		try {
			workbook=new XSSFWorkbook();
			ByteArrayOutputStream out = new ByteArrayOutputStream();
			Sheet sheet=workbook.createSheet("finacial data");
			Row headeRrow=	sheet.createRow(0);
            String[] headers = {"ID", "Account", "Business Unit", "Currency", "Scenario", "Year", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"};
			for(int i =0;i<headers.length;i++) {
            Cell headerCell=	headeRrow.createCell(i);
            headerCell.setCellValue(headers[i]);
			}
			//write financila data to rows
            int rowNumber = 1; // Starting from 1 as 0 is header row

			for(FinancialRecords financialRecords:financialRecordsList) {
				
				 Row row = sheet.createRow(rowNumber++);

	                row.createCell(0).setCellValue(financialRecords.getId());
	                row.createCell(1).setCellValue(financialRecords.getAccount());
	                row.createCell(2).setCellValue(financialRecords.getBusinessUnit());
	                row.createCell(3).setCellValue(financialRecords.getCurrency());
	                row.createCell(4).setCellValue(financialRecords.getScenario());
	                row.createCell(5).setCellValue(financialRecords.getYear());;

				
				
				BigDecimal[]months=	financialRecords.getMonths();
				for (int i = 0; i < months.length; i++) {
				    row.createCell(6 + i).setCellValue(months[i] != null ? months[i].doubleValue() : 0.0);
				}
			}
			workbook.write(out);
		      return new ByteArrayInputStream(out.toByteArray());

           
		}
		catch (IOException e) {
			throw new RuntimeException("Error while writing financial records to excel:"+e.getMessage());
		}
		finally {
			workbook.close();
		}
	}


	public static List<FinancialRecords> excelToFinacialRecords(InputStream inputStream) {
		try {
			System.err.println("sheet name"+SHEET);
		Workbook workbook=new XSSFWorkbook(inputStream);
		System.err.println(workbook.getSheetName(0));
		Sheet sheet=	workbook.getSheet(SHEET);
		Iterator<Row>rows=  sheet.iterator();
		List<FinancialRecords>financialRecordsList=new  ArrayList();
		int rowNumber=0;
		
		while(rows.hasNext()) {
		Row currentRow=	rows.next();
		if(rowNumber==0) {
			rowNumber++;
			continue;
		}
	Iterator<Cell>cellsInRow=	currentRow.iterator();
		FinancialRecords financialRecords=new FinancialRecords();
		 BigDecimal[] monthValues = new BigDecimal[12];
		int cellIndex=0;
		
		while(cellsInRow.hasNext()) {
			
		Cell currentCell=	cellsInRow.next();
		
		switch(cellIndex) {
	
			
		case 0:
			financialRecords.setAccount(currentCell.getStringCellValue());
			break;
		case 1:
			financialRecords.setBusinessUnit(currentCell.getStringCellValue());
			break;
		case 2:
			financialRecords.setCurrency(currentCell.getStringCellValue());
			break;
		case 3:
			financialRecords.setYear((int)currentCell.getNumericCellValue());
			break;
		case 4:
			financialRecords.setScenario(currentCell.getStringCellValue());
			break;
			default:
				if(cellIndex>=4&&cellIndex<=15) {
					monthValues[cellIndex-5]=BigDecimal.valueOf(currentCell.getNumericCellValue());
				}
				break;
				
			
		}
		cellIndex++;
		
		}
		financialRecords.setMonths(monthValues);
		financialRecordsList.add(financialRecords);
		
		
		}
		workbook.close();
		return financialRecordsList;
		}
		catch (IOException e) {
			throw new RuntimeException("Fail to parse excel file:"+e.getMessage());
}
		
		
	}

}
