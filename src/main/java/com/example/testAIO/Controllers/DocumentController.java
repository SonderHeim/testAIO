package com.example.testAIO.Controllers;
import org.apache.commons.io.FileUtils;
import org.apache.poi.xwpf.usermodel.*;
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


@RestController
public class DocumentController {

    @GetMapping(value = "/documentRep", produces = MediaType.APPLICATION_OCTET_STREAM_VALUE)
    public ResponseEntity<byte[]> generateDocument() throws Exception {
        // Считываем список людей
        List<Person> personList = getPersonList();

        // Открываем шаблон документа Word
        FileInputStream templateInputStream = new FileInputStream("E:/retire/testAIO/document.docx");
        XWPFDocument document = new XWPFDocument(templateInputStream);

        // Заменяем метки на данные из списка
        replacePlaceholder(document, "##FIRST_NAME##", personList.get(0).getFirstName());
        replacePlaceholder(document, "##LAST_NAME##", personList.get(0).getLastName());
        replacePlaceholder(document, "##AGE##", String.valueOf(personList.get(0).getAge()));

        // Сохраняем изменения в документе
        FileOutputStream out = new FileOutputStream("E:/retire/testAIO/output.docx");
        document.write(out);
        out.close();

        // Отправляем документ клиенту
        File file = new File("output.docx");
        byte[] documentBytes = FileUtils.readFileToByteArray(file);
        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
        headers.setContentDispositionFormData("attachment", file.getName());
        headers.setContentLength(documentBytes.length);
        return new ResponseEntity<>(documentBytes, headers, HttpStatus.OK);
    }

    // Метод для замены меток на данные из списка
    private void replacePlaceholder(XWPFDocument document, String placeholder, String value) {
        for (XWPFParagraph p : document.getParagraphs()) {
            List<XWPFRun> runs = p.getRuns();
            if (runs != null) {
                for (XWPFRun r : runs) {
                    String text = r.getText(0);
                    if (text != null && text.contains(placeholder)) {
                        text = text.replace(placeholder, value);
                        r.setText(text, 0);
                    }
                }
            }
        }
        for (XWPFTable tbl : document.getTables()) {
            for (XWPFTableRow row : tbl.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (XWPFParagraph p : cell.getParagraphs()) {
                        for (XWPFRun r : p.getRuns()) {
                            String text = r.getText(0);
                            if (text != null && text.contains(placeholder)) {
                                text = text.replace(placeholder, value);
                                r.setText(text, 0);
                            }
                        }
                    }
                }
            }
        }
    }

    // Метод для получения списка людей
    private List<Person> getPersonList() {
        List<Person> personList = new ArrayList<>();
        personList.add(new Person("Иван", "Иванов", 33));
        personList.add(new Person("Петр", "Петров", 28));
        personList.add(new Person("Сергей", "Сергеев", 45));
        return personList;
    }

    // Класс для хранения информации о человеке
    private static class Person {
        private String firstName;
        private String lastName;
        private int age;

        public Person(String firstName, String lastName, int age) {
            this.firstName = firstName;
            this.lastName = lastName;
            this.age = age;
        }

        public String getFirstName() {
            return firstName;
        }

        public void setFirstName(String firstName) {
            this.firstName = firstName;
        }

        public String getLastName() {
            return lastName;
        }

        public void setLastName(String lastName) {
            this.lastName = lastName;
        }

        public int getAge() {
            return age;
        }

        public void setAge(int age) {
            this.age = age;
        }
    }
}