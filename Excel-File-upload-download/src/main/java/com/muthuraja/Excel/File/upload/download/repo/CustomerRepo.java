package com.muthuraja.Excel.File.upload.download.repo;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

import com.muthuraja.Excel.File.upload.download.model.Customer;

@Repository
public interface CustomerRepo extends JpaRepository<Customer, Long>{

}
