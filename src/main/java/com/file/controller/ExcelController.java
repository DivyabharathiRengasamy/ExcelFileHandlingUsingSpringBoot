package com.file.controller;

import java.io.IOException;
import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import com.file.helper.ExcelHelper;
import com.file.message.ResponseMessage;
import com.file.model.FinancialRecords;
import com.file.service.ExcelService;

import jakarta.annotation.Resource;

@RestController
@RequestMapping("/api/excel")
public class ExcelController {
	
	@Autowired
	ExcelService excelService;
	@PostMapping("/upload")
	public ResponseEntity<ResponseMessage>uploadFile(@RequestParam ("file") 
	MultipartFile file){
		String message="";
		
		if(ExcelHelper.hasExcelFormat(file)) {
			try {
				excelService.save(file);
				message="uploaded the fail successfully:"+file.getOriginalFilename();
		        return ResponseEntity.status(HttpStatus.OK).body(new ResponseMessage(message));
			} catch (IOException e) {
				message="fail Could not br successfully uploaded :"+file.getOriginalFilename()+"!";
				e.printStackTrace();
				return ResponseEntity.status(HttpStatus.EXPECTATION_FAILED).body(new ResponseMessage(message));
			}
		}
			 message = "Please upload an excel file!";
			    return ResponseEntity.status(HttpStatus.BAD_REQUEST).body(new ResponseMessage(message));
			  }
	
	@GetMapping("/getAllFinancialRecords")
	public ResponseEntity<List<FinancialRecords>> getAllFinancialRecords() {
		try {
			
		
		List <FinancialRecords>financialRecordsList=excelService.getAllFinancialRecords();
		if(financialRecordsList.isEmpty()) {
			return new ResponseEntity<>(HttpStatus.NO_CONTENT);
		}
		
	      return new ResponseEntity<>(financialRecordsList, HttpStatus.OK);
		}
	
		
		catch (Exception e) {
			return new ResponseEntity<>(null,HttpStatus.INTERNAL_SERVER_ERROR);
		}
		
		
	}
	
	@GetMapping("/download")
	public ResponseEntity<InputStreamResource> downloadFile() throws IOException {
		String filename="financialData.xlsx";
	
	    InputStreamResource inputStreamResource=new InputStreamResource(excelService.load());
		
	    return ResponseEntity.ok().header(HttpHeaders.CONTENT_DISPOSITION, "filename="+filename)
	    		.contentType(MediaType.parseMediaType("application/vnd.ms-excel"))
	    		.body(inputStreamResource);
	    		
	}
}
	
	
	
	

