package com.tpx.amper.charles;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.CommandLineRunner;

import java.util.Iterator;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

@SpringBootApplication
public class ServiceLineDataToolApplication implements CommandLineRunner {
	
	private static final Logger LOGGER = LoggerFactory.getLogger(ServiceLineDataToolApplication.class);

	public static void main(String[] args) {
		SpringApplication.run(com.tpx.amper.charles.ServiceLineDataToolApplication.class, args);
	}
	
    @Override
    public void run(String... args) throws Exception {
        String spreadsheetPath = "sat_names.xlsx";
        BufferedWriter writer = null;
        Workbook workbook = null;
        FileInputStream fileInputStream = null;
        try {
            // Load spreadsheet
            fileInputStream = new FileInputStream(spreadsheetPath);
            workbook = new XSSFWorkbook(fileInputStream);
            Sheet sheet = workbook.getSheetAt(0); // Assuming the data is in the first sheet
            writer = new BufferedWriter(new FileWriter("output.sql"));
                // Iterate through rows in the spreadsheet 
            	int count = 1;
            	
            	writer.write("SET DEFINE OFF;");
            	writer.newLine();
            	
            	Iterator<Row> ite = sheet.iterator();
            	ite.next(); // don't include the header

            	while (ite.hasNext()) {
            		
            		Row row = (Row) ite.next();
            		
            		Cell cell1 = row.getCell(0);
                    Cell cell2 = row.getCell(1);
                    Cell cell3 = row.getCell(2);
                    Cell cell4 = row.getCell(4);
                    
                    LOGGER.info("Writing insert scripts into the file");
                    
                    writer.write("INSERT INTO SALT.SAT_LOOKUP (ENUM_TYPE, KEY, VALUE, VALUE2, VALUE3, ACTIVE, CREATEDBY, CREATEDDATE, MODIFIEDBY, MODIFIEDDATE, VALUE4) " + 
                    		"VALUES " + 
                    		"('ServiceLine', '" + count +"', '" + cell1.getStringCellValue().trim() + "', '" + cell2.getStringCellValue().trim() + "','"+cell3.getStringCellValue().trim()+"', '1', 'charles.amper', CURRENT_TIMESTAMP, 'charles.amper', CURRENT_TIMESTAMP, '"+cell4.getStringCellValue().trim()+"');");
                    LOGGER.info("INSERT INTO SALT.SAT_LOOKUP (ENUM_TYPE, KEY, VALUE, VALUE2, VALUE3, ACTIVE, CREATEDBY, CREATEDDATE, MODIFIEDBY, MODIFIEDDATE, VALUE4) " + 
                    		"VALUES " + 
                    		"('ServiceLine', '" + count +"', '" + cell1.getStringCellValue().trim() + "', '" + cell2.getStringCellValue().trim() + "','"+cell3.getStringCellValue().trim()+"', '1', 'charles.amper', CURRENT_TIMESTAMP, 'charles.amper', CURRENT_TIMESTAMP, '"+cell4.getStringCellValue().trim()+"');");
                    writer.newLine();
                    count++;
                    
            	}
            	LOGGER.info("Completed Successfully!");
      
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
        	writer.close();
            workbook.close();
        }

    }

}