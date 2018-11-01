package com.mikerusoft.excel2word;

import org.apache.poi.xwpf.usermodel.*;
import org.springframework.stereotype.Component;

import java.io.*;
import java.util.List;

@Component
public class WordFileDataOutputer implements DataOutputer {

    @Override
    public void output(List<String> content, String fileName) {
        if (new File(fileName).exists()) {
            throw new RuntimeException("Such file already exists " + fileName);
        }
        try(FileOutputStream out = new FileOutputStream(fileName)) {
            XWPFDocument document = new XWPFDocument();

            for(String x : content) {
                XWPFParagraph paragraph = document.createParagraph();
                paragraph.setWordWrapped(true);
                XWPFRun run = paragraph.createRun();
                run.setText(x);
            }
            document.write(out);
            document.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
