package cares.cwds.salesforce.common.utilities;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

public class PoiWriteExcel {

	 private PoiWriteExcel(){
	    	
	    	/*This method is initially left blank for now*/
	    }
	 
	private static final Logger logger =LoggerFactory.getLogger(PoiWriteExcel.class.getName());
	
	public static void writeExcelData(String filePath, String sheetName, String[] fieldNames, String fieldValue) {
	    try (FileInputStream fis = new FileInputStream(filePath);
	         XSSFWorkbook workbook = new XSSFWorkbook(fis)) {
 
	        Sheet sheet = workbook.getSheet(sheetName);
	        Row headerRow = sheet.getRow(0);
	        Map<String, Integer> colByName = IntStream.range(0, headerRow.getLastCellNum())
	            .boxed()
	            .collect(Collectors.toMap(
	                i -> new DataFormatter().formatCellValue(headerRow.getCell(i)),
	                i -> i,
	                (v1, v2) -> v1,
	                LinkedHashMap::new
	            ));
 
	        Row newRow = sheet.createRow(sheet.getLastRowNum() + 1);
	        String currentDateTime = LocalDateTime.now().format(DateTimeFormatter.ofPattern("dd/MM/yy HH:mm"));
 
	        colByName.forEach((columnName, columnIndex) -> {
	            if (columnName.equals(fieldNames[0])) {
	                newRow.createCell(columnIndex).setCellValue(fieldValue);
	            } else if (columnName.equals(fieldNames[1])) {
	                newRow.createCell(columnIndex).setCellValue(currentDateTime);
	            }
	        });
 
	        try (FileOutputStream fos = new FileOutputStream(filePath)) {
	            workbook.write(fos);
	        }
	    } catch (IOException e) {
	        logger.info("Error writing Excel data");
	    }
	}
 
 
	    public static String readScreeningId(String filePath, String sheetName, String[] fieldNames, String scId){
	        try (FileInputStream fis = new FileInputStream(filePath);
	             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {
	            
	            Sheet sheet = workbook.getSheet(sheetName);
	            Row headerRow = sheet.getRow(0);
	            Map<String, Integer> colByName = IntStream.range(0, headerRow.getLastCellNum())
	                .boxed()
	                .collect(Collectors.toMap(
	                    i -> new DataFormatter().formatCellValue(headerRow.getCell(i)),
	                    i -> i
	                ));
 
	            DataFormatter formatter = new DataFormatter();
	            String formattedDateTime = LocalDateTime.now().format(DateTimeFormatter.ofPattern("dd/MM/yy HH:mm"));
 
	            for (Row row : sheet) {
	                if (row.getRowNum() == 0) continue; // Skip header row
	                
	                String scDataValue = formatter.formatCellValue(row.getCell(colByName.get(fieldNames[0])));
	                if (scDataValue.isEmpty()) {
	                    String screeningId = formatter.formatCellValue(row.getCell(colByName.get(scId)));
	                    row.createCell(colByName.get(fieldNames[0])).setCellValue("USED");
	                    row.createCell(colByName.get(fieldNames[1])).setCellValue(formattedDateTime);
	                    
	                    try (FileOutputStream fos = new FileOutputStream(filePath)) {
	                        workbook.write(fos);
	                    }
	                    
	                    return screeningId;
	                }
	            }
	        } catch (Exception e) {
	        	logger.info("Error reading screening ID");
	        }
	        return "";
	    }
	}

