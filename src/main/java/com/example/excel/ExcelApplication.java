package com.example.excel;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import tk.mybatis.spring.annotation.MapperScan;

@SpringBootApplication
@MapperScan("com.example.excel.mapper")
public class ExcelApplication {
    public static void main(String[] args) {
        SpringApplication.run(ExcelApplication.class, args);
    }

}
