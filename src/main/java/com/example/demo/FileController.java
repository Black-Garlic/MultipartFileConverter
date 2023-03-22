package com.example.demo;

import kr.dogfoot.hwplib.object.HWPFile;
import kr.dogfoot.hwplib.reader.HWPReader;
import kr.dogfoot.hwplib.tool.textextractor.TextExtractMethod;
import kr.dogfoot.hwplib.tool.textextractor.TextExtractor;
import org.apache.commons.io.IOUtils;
import org.apache.poi.hssf.extractor.ExcelExtractor;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xssf.extractor.XSSFExcelExtractor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.nio.charset.Charset;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

@RestController
public class FileController {

    @PostMapping("/file")
    public void uploadDocument(@RequestParam(value = "file") MultipartFile multipartFile) throws Exception {

        String parseData = "";

        String imageExtension = "";

        String[] imageExtensionSplit = multipartFile.getOriginalFilename().split("\\.");
        imageExtension = imageExtensionSplit[imageExtensionSplit.length - 1];

        switch (imageExtension) {
            case "docx" :
                parseData = parseDocx(multipartFile);
                break;
            case "doc" :
                parseData = parseDoc(multipartFile);
                break;
            case "xlsx" :
                parseData = parseXlsx(multipartFile);
                break;
            case "xls" :
                parseData = parseXls(multipartFile);
                break;
            case "csv" :
                parseData = parseCsv(multipartFile);
                break;
            case "hwp" :
                parseData = parseHwp(multipartFile);
                break;
        }

        System.out.println(parseData);
    }

    public String parseDocx(MultipartFile multipartFile) throws Exception {
        StringBuilder sb = new StringBuilder();

        try {
            XWPFDocument xwpfDocument = new XWPFDocument(OPCPackage.open(multipartFile.getInputStream()));

            XWPFWordExtractor xwpfWordExtractor = new XWPFWordExtractor(xwpfDocument);

            sb.append(xwpfWordExtractor.getText());
        } catch(Exception e) {
            throw new Exception(e);
        }

        return sb.toString();
    }

    public String parseDoc(MultipartFile multipartFile) throws Exception {
        StringBuilder sb = new StringBuilder();

        try {
            POIFSFileSystem poiFS = new POIFSFileSystem(multipartFile.getInputStream());

            HWPFDocument hwpfDocument = new HWPFDocument(poiFS);

            WordExtractor wordExtractor = new WordExtractor(hwpfDocument);

            String[] paragraphs = wordExtractor.getParagraphText();

            for (String paragraph : paragraphs) {
                sb.append(paragraph);
            }
        } catch (Exception e) {
            throw new Exception(e);
        }

        return sb.toString();
    }

    public String parseXlsx(MultipartFile multipartFile) throws Exception {
        StringBuilder sb = new StringBuilder();

        try {
            XSSFWorkbook xssfWorkbook = new XSSFWorkbook(multipartFile.getInputStream());

            XSSFExcelExtractor xssfExcelExtractor = new XSSFExcelExtractor(xssfWorkbook);
            xssfExcelExtractor.setFormulasNotResults(true);
            xssfExcelExtractor.setIncludeSheetNames(false);

            sb.append(xssfExcelExtractor.getText());
        } catch (Exception e) {
            throw new Exception(e);
        }

        return sb.toString();
    }

    public String parseXls(MultipartFile multipartFile) throws Exception {
        StringBuilder sb = new StringBuilder();

        try {
            HSSFWorkbook hssfWorkbook = new HSSFWorkbook(multipartFile.getInputStream());

            ExcelExtractor excelExtractor = new ExcelExtractor(hssfWorkbook);

            sb.append(excelExtractor.getText());
        } catch (Exception e) {
            throw new Exception(e);
        }

        return sb.toString();
    }

    public String parseCsv(MultipartFile multipartFile) throws Exception {
        byte[] fileByteArray = IOUtils.toByteArray(multipartFile.getInputStream());
        StringBuilder sb = new StringBuilder();

        try {
            BufferedReader br = new BufferedReader(new StringReader(new String(fileByteArray)));
            Charset.forName("UTF-8");
            String line = "";

            while ((line = br.readLine()) != null) {
                String[] token = line.split(",");
                List<String> tempList = new ArrayList<String>(Arrays.asList(token));

                for (String cell : tempList) {
                    sb.append(cell).append("\t");
                }
                sb.append("\n");
            }
        }catch (Exception e){
            throw new Exception(e);
        }


        return sb.toString();
    }

    public String parseHwp(MultipartFile multipartFile) throws Exception {
        StringBuilder sb = new StringBuilder();

        try {
            HWPFile hwpFile = HWPReader.fromInputStream(multipartFile.getInputStream());

            sb.append(TextExtractor.extract(hwpFile, TextExtractMethod.InsertControlTextBetweenParagraphText));
        } catch (Exception e) {
            throw new Exception(e);
        }

        return sb.toString();
    }
}
