package com.muthuraja.Excel.File.upload.download.controller;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import com.muthuraja.Excel.File.upload.download.model.Customer;
import com.muthuraja.Excel.File.upload.download.repo.CustomerRepo;

@Controller
@RequestMapping("/file")
public class CustomerController {

	@Autowired
	private CustomerRepo customerRepo;
	
	@GetMapping("/")
    public String index() {
        return "multipartFile/Excel_Upload_Download.html";
    }
    
	
	 @PostMapping("/uploadFile")
	    public String uploadMultipartFile(@RequestParam("uploadfile") MultipartFile file, Model model) {
			try {
				store(file);
				model.addAttribute("message", "File uploaded successfully!");
			} catch (Exception e) {
				model.addAttribute("message", "Fail! -> uploaded filename: " + file.getOriginalFilename());
			}
	        return "multipartFile/Excel_Upload_Download.html";
	    }
	 
	 @GetMapping("/downloadFile")
		public ResponseEntity<InputStreamResource> downloadFile() {
			
			HttpHeaders headers = new HttpHeaders();
	        headers.add("Content-Disposition", "attachment; filename=customers.xlsx");
			
			return ResponseEntity
	                .ok()
	                .headers(headers)
	                .body(new InputStreamResource(loadFile()));	
		}
	 
	 public void store(MultipartFile file){
			try {
				List<Customer> lstCustomers = parseExcelFile(file.getInputStream());
	    		customerRepo.saveAll(lstCustomers);
	        } catch (IOException e) {
	        	throw new RuntimeException("FAIL! -> message = " + e.getMessage());
	        }
		}
	 
	 public static List<Customer> parseExcelFile(InputStream is) {
			try {
	    		Workbook workbook = new XSSFWorkbook(is);
	     
	    		Sheet sheet = workbook.getSheet("Customers");
	    		Iterator<Row> rows = sheet.iterator();
	    		
	    		List<Customer> lstCustomers = new ArrayList<Customer>();
	    		
	    		int rowNumber = 0;
	    		while (rows.hasNext()) {
	    			Row currentRow = rows.next();
	    			
	    			// skip header
	    			if(rowNumber == 0) {
	    				rowNumber++;
	    				continue;
	    			}
	    			
	    			Iterator<Cell> cellsInRow = currentRow.iterator();

	    			Customer cust = new Customer();
	    			
	    			int cellIndex = 0;
	    			while (cellsInRow.hasNext()) {
	    				Cell currentCell = cellsInRow.next();
	    				
	    				if(cellIndex==0) { // ID
	    					cust.setId((long) currentCell.getNumericCellValue());
	    				} else if(cellIndex==1) { // Name
	    					cust.setName(currentCell.getStringCellValue());
	    				} else if(cellIndex==2) { // Address
	    					cust.setAddress(currentCell.getStringCellValue());
	    				} else if(cellIndex==3) { // Age
	    					cust.setAge((int) currentCell.getNumericCellValue());
	    				}
	    				
	    				cellIndex++;
	    			}
	    			
	    			lstCustomers.add(cust);
	    		}
	    		
	    		// Close WorkBook
	    		workbook.close();
	    		
	    		return lstCustomers;
	        } catch (IOException e) {
	        	throw new RuntimeException("FAIL! -> message = " + e.getMessage());
	        }
		}
	 
	  public ByteArrayInputStream loadFile() {
	    	List<Customer> customers = (List<Customer>) customerRepo.findAll();
	    	
	    	try {
	    		ByteArrayInputStream in = customersToExcel(customers);
	    		return in;
			} catch (IOException e) {}
	    	
	        return null;
	    }
	  public static ByteArrayInputStream customersToExcel(List<Customer> customers) throws IOException {
			String[] COLUMNs = {"Id", "Name", "Address", "Age"};
			try(
					Workbook workbook = new XSSFWorkbook();
					ByteArrayOutputStream out = new ByteArrayOutputStream();
			){
				CreationHelper createHelper = workbook.getCreationHelper();
		 
				Sheet sheet = workbook.createSheet("Customers");
		 
				Font headerFont = workbook.createFont();
				headerFont.setBold(true);
				headerFont.setColor(IndexedColors.BLUE.getIndex());
		 
				CellStyle headerCellStyle = workbook.createCellStyle();
				headerCellStyle.setFont(headerFont);
		 
				// Row for Header
				Row headerRow = sheet.createRow(0);
		 
				// Header
				for (int col = 0; col < COLUMNs.length; col++) {
					Cell cell = headerRow.createCell(col);
					cell.setCellValue(COLUMNs[col]);
					cell.setCellStyle(headerCellStyle);
				}
		 
				// CellStyle for Age
				CellStyle ageCellStyle = workbook.createCellStyle();
				ageCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("#"));
		 
				int rowIdx = 1;
				for (Customer customer : customers) {
					Row row = sheet.createRow(rowIdx++);
		 
					row.createCell(0).setCellValue(customer.getId());
					row.createCell(1).setCellValue(customer.getName());
					row.createCell(2).setCellValue(customer.getAddress());
		 
					Cell ageCell = row.createCell(3);
					ageCell.setCellValue(customer.getAge());
					ageCell.setCellStyle(ageCellStyle);
				}
		 
				workbook.write(out);
				return new ByteArrayInputStream(out.toByteArray());
			}
		}
}
