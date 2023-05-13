package com.example.testAIO.Controllers;

import com.example.testAIO.models.Person;
import com.example.testAIO.repositories.PersonRepository;
import org.apache.commons.io.FileUtils;
import org.apache.poi.xwpf.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.Resource;
import org.springframework.core.io.UrlResource;
import org.springframework.http.*;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.springframework.core.io.ClassPathResource;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
@Controller
public class PersonController {

    @Autowired
    private PersonRepository personRepository;

    @GetMapping("/generate-doc")
    public ResponseEntity<Resource> generateDoc() throws IOException {

        // Получаем данные из базы данных
        List<Person> people = personRepository.findAll();

        // Читаем шаблон документа
        ClassLoader classLoader = getClass().getClassLoader();
        File file = new File("E:/retire/testAIO/document.docx");
        FileInputStream fis = new FileInputStream(file);
        XWPFDocument document = new XWPFDocument(fis);
        fis.close();

        // Заполняем шаблон данными
        for (XWPFParagraph paragraph : document.getParagraphs()) {
            List<XWPFRun> runs = paragraph.getRuns();
            for (XWPFRun run : runs) {
                String text = run.getText(0);
                if (text != null && text.contains("<<Имя>>")) {
                    text = text.replace("<<Имя>>", people.get(0).getFirstName());
                    run.setText(text, 0);
                }
                if (text != null && text.contains("<<Фамилия>>")) {
                    text = text.replace("<<Фамилия>>", people.get(0).getLastName());
                    run.setText(text, 0);
                }
                if (text != null && text.contains("<<Возраст>>")) {
                    text = text.replace("<<Возраст>>", Integer.toString(people.get(0).getAge()));
                    run.setText(text, 0);
                }
            }
        }

        // Сохраняем заполненный документ во временный файл
        File tempFile = new File("E:/retire/testAIO/result.docx");
        FileOutputStream fos = new FileOutputStream(tempFile);
        document.write(fos);
        fos.close();
        document.close();

        // Отдаем файл в ответе
        HttpHeaders headers = new HttpHeaders();
        headers.add(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=document.docx");
        headers.add(HttpHeaders.CONTENT_TYPE, "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
        Resource resource = new UrlResource(tempFile.toURI());
        return ResponseEntity.ok()
                .headers(headers)
                .body(resource);
    }
}

