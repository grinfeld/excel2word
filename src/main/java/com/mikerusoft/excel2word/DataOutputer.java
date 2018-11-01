package com.mikerusoft.excel2word;

import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.util.List;

public interface DataOutputer {
    void output(List<String> content, String fileName);
}
