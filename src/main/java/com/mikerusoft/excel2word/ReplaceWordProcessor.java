package com.mikerusoft.excel2word;

import com.mikerusoft.excel2word.props.word.WordProperties;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

import java.io.*;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.*;
import java.util.stream.Collectors;

@Component
public class ReplaceWordProcessor implements WordProcessor<List<List<String>>> {

    private List<String> files;

    @Autowired
    public ReplaceWordProcessor(WordProperties properties) {
        this.files = Arrays.asList(properties.getTemplates());
    }

    @Override
    public List<List<String>> processData(List<Map<String, String>> data) {
        return files.stream().map(ReplaceWordProcessor::getFileFromName)
            .map(c -> getParsedContent(c, data)).collect(Collectors.toList());
    }

    private static List<String> getParsedContent(Path path, List<Map<String, String>> data) {

        try (XWPFDocument document = new XWPFDocument(new FileInputStream(path.toFile()))) {
            return Optional.ofNullable(document.getParagraphs()).orElseGet(ArrayList::new)
                    .stream().map(paragraph -> replaceData(paragraph.getText(), data))
                    .filter(Objects::nonNull).collect(Collectors.toList());
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    private static XWPFParagraph getXwpfParagraph(XWPFDocument document, String t) {
        try {
            XWPFParagraph para = document.createParagraph();
            XWPFRun para2Run = para.createRun();
            para2Run.setText(t);
            return para;
        } catch (Exception e) {
            return null;
        }
    }

    private static String replaceData(String line, List<Map<String, String>> data) {
        String newLine = line;
        for (Map<String, String> fields : data) {
            for (Map.Entry<String, String> entry : fields.entrySet()) {
                newLine = newLine.replaceAll("\\{\\{" + entry.getKey() + "}}", entry.getValue());
            }
        }
        return newLine;
    }

    private static Path getFileFromName(String fileName) {
        File file = new File(fileName);
        if (file.exists())
            return Paths.get(file.getAbsolutePath());

        file = new File(ClassLoader.getSystemResource(fileName).getFile());
        if (file.exists())
            return Paths.get(file.getAbsolutePath());

        throw new RuntimeException("Failed to find file " + fileName);
    }


}
