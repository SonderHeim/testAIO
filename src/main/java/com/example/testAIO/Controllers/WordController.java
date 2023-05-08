package com.example.testAIO.Controllers;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;
import java.awt.Font;
import org.apache.poi.xwpf.usermodel.XWPFStyle;
import java.time.LocalDate;
import java.util.HashMap;
import java.util.Map;
import java.io.File;
import java.io.FileOutputStream;
import java.nio.file.Files;

import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;


import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.List;

@RestController
public class WordController {

    @GetMapping(value = "/generate-word-doc", produces = MediaType.APPLICATION_OCTET_STREAM_VALUE)
    public ResponseEntity<byte[]> generateWordDocument() throws IOException {
        // Создаем новый документ Word
        XWPFDocument document = new XWPFDocument();

        // Создаем новый параграф и добавляем текст в документ
        XWPFParagraph firstPar = document.createParagraph();
        firstPar.setSpacingBefore(0);
        XWPFRun firstRun = firstPar.createRun();
        firstPar.setIndentationFirstLine(200); // устанавливаем отступ первой строки в 1 см
        firstRun.setText("Привет, мир!");
        firstRun.setBold(true);
        firstRun.addCarriageReturn();


//        firstRun.addBreak();
//
//        XWPFParagraph secondPar = document.createParagraph();
//        secondPar.setSpacingBefore(0);
//        XWPFRun secondRun = secondPar.createRun();
//        secondPar.setIndentationFirstLine(200); // устанавливаем отступ первой строки в 1 см
//        secondRun.setText("Итак, предположим, у вас на руках есть (ненужный) файл docx. Преобразуем его в файл zip (осторожно, обратное преобразование путем переименования zip -> docx может сделать файл недоступным для вашего редактора(!)), в получившемся архиве откроем папку word, а в ней — файл document.xml. Перед нами xml-представление word-файла, которое также можно было бы получить через Apache POI, с меньшими трудностями.");

        for (XWPFParagraph p : document.getParagraphs()) {
            for (XWPFRun r : p.getRuns()) {
                r.setItalic(false);
                r.setUnderline(UnderlinePatterns.NONE);
                r.setFontFamily("Times New Roman");
                r.setFontSize(14);
            }
        }

        // Создаем массив байт для записи документа в него
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        document.write(outputStream);

        // Задаем заголовок с именем файла
        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
        headers.setContentDispositionFormData("attachment", "example.docx");

        // Возвращаем файл в качестве ответа на запрос клиента
        return ResponseEntity.ok()
                .headers(headers)
                .body(outputStream.toByteArray());
    }
    @GetMapping("/generate-word")
    public ResponseEntity<byte[]> generateWord() throws IOException {

        // Создание документа Word
        XWPFDocument document = new XWPFDocument();

        // Создание таблицы
        XWPFTable table = document.createTable();

        // Создание заголовка таблицы
        XWPFTableRow header = table.getRow(0);
        header.getCell(0).setText("Имя");
        header.createCell().setText("Фамилия");
        header.createCell().setText("Возраст");

        // Создание списка объектов Person
        List<Person> personList = createPersonList();

        // Заполнение таблицы данными из списка
        for (Person person : personList) {
            XWPFTableRow row = table.createRow();
            row.getCell(0).setText(person.getFirstName());
            row.getCell(1).setText(person.getLastName());
            row.getCell(2).setText(Integer.toString(person.getAge()));
        }

        XWPFParagraph firstPar = document.createParagraph();
        firstPar.setSpacingBefore(0);
        XWPFRun firstRun = firstPar.createRun();
        firstPar.setIndentationFirstLine(200); // устанавливаем отступ первой строки в 1 см
        firstRun.setText("${title}");

        // Замена меток в документе на значения
        Map<String, String> replaceMap = new HashMap<>();
        replaceMap.put("${title}", "Список людей");
        replaceMap.put("${date}", LocalDate.now().toString());
        replaceInDocument(document, replaceMap);

        // Сохранение документа в файл
        File file = new File("persons.docx");
        FileOutputStream fos = new FileOutputStream(file);
        document.write(fos);

        // Отправка документа в ответе на запрос
        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
        headers.setContentDispositionFormData("attachment", "persons.docx");
        byte[] bytes = Files.readAllBytes(file.toPath());
        return new ResponseEntity<>(bytes, headers, HttpStatus.OK);
    }
    private void replaceInDocument(XWPFDocument document, Map<String, String> replaceMap) {
        for (XWPFParagraph p : document.getParagraphs()) {
            List<XWPFRun> runs = p.getRuns();
            if (runs != null) {
                for (XWPFRun r : runs) {
                    String text = r.getText(0);
                    if (text != null) {
                        for (Map.Entry<String, String> entry : replaceMap.entrySet()) {
                            if (text.contains(entry.getKey())) {
                                text = text.replace(entry.getKey(), entry.getValue());
                                r.setText(text, 0);
                            }
                        }
                    }
                }
            }
        }
    }



    public class Person {
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

    public List<Person> createPersonList() {
        List<Person> personList = new ArrayList<>();
        personList.add(new Person("John", "Doe", 30));
        personList.add(new Person("Jane", "Smith", 25));
        personList.add(new Person("Bob", "Johnson", 40));
        return personList;
    }



}
