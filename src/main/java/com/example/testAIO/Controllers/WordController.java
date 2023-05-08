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


import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.math.BigInteger;

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
}
