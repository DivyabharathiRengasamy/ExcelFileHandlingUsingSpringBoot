package com.file.service;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.util.*;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.InputStreamSource;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import com.file.helper.ExcelHelper;
import com.file.model.FinancialRecords;
import com.file.repository.FileRepository;

@Service
public class ExcelService {
@Autowired
	FileRepository fileRepository;
	public void save(MultipartFile file) throws IOException {
		try {
			System.err.println(file.getOriginalFilename());
			System.err.println(file.getInputStream());

		List<FinancialRecords>list=	ExcelHelper.excelToFinacialRecords(file.getInputStream());
		fileRepository.saveAll(list);
		}
		catch (IOException e) {
			throw new RuntimeException("Failed to store data:"+e.getMessage());
		}
	}
	public List<FinancialRecords> getAllFinancialRecords() throws IOException {
	List<FinancialRecords>financialRecordsList=	fileRepository.findAll();
	
		return financialRecordsList;
	}
	public ByteArrayInputStream load() throws IOException {
		List<FinancialRecords>financialRecordsList=	fileRepository.findAll();
			ByteArrayInputStream arrayInputStream=ExcelHelper.finanacialRecorsdToExcel(financialRecordsList);
			return arrayInputStream;
	}
	
	

	

}
