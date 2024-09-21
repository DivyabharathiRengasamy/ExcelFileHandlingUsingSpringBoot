package com.file.repository;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

import com.file.model.FinancialRecords;
@Repository
public interface FileRepository extends JpaRepository<FinancialRecords, Long>{

}
