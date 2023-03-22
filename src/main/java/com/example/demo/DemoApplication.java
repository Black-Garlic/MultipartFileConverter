package com.example.demo;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.FileInputStream;

@SpringBootApplication
public class DemoApplication {

	public static void main(String[] args) {
		SpringApplication.run(DemoApplication.class, args);
	}

	public static void test () {
		try {
			FileInputStream fis = new FileInputStream("/Users/haneul/IdeaProjects/DocumentReader/test.docx");
			XWPFDocument file   = new XWPFDocument(OPCPackage.open(fis));
			XWPFWordExtractor ext = new XWPFWordExtractor(file);

			System.out.println("===== docx text extractor ======");
			System.out.println(ext.getText());
		}catch(Exception e) {
			System.out.println(e);
		}
	}

}
