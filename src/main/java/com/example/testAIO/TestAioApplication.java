package com.example.testAIO;
import org.docx4j.Docx4J;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.File;

@SpringBootApplication
public class TestAioApplication {

	public static void main(String[] args) throws Exception {
		SpringApplication.run(TestAioApplication.class, args);
	}
}
