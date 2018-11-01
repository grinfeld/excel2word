package com.mikerusoft.excel2word;

import com.mikerusoft.excel2word.props.excel.ExcelProperties;
import com.mikerusoft.excel2word.utils.ImmutablePair;
import com.mikerusoft.excel2word.utils.Streams;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Component;

import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.stream.Collectors;

@Slf4j
@Component
public class ExcelDataReader implements DataReader {

    private String fileName;
    private String sheetName;

    public ExcelDataReader(ExcelProperties properties) {
        this.fileName = properties.getFileOptional().map(ExcelProperties.File::getName).orElse(null);
        this.sheetName = properties.getSheetOptional().map(ExcelProperties.Sheet::getName).orElse(null);
    }

    @Override
    public List<Map<String, String>> readData() {
        try {
            Workbook workbook = WorkbookFactory.create(new File(fileName));
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

            Sheet sheet = workbook.getSheet(sheetName);
            if (sheet == null)
                throw new RuntimeException("Failed to find sheet with name " + sheetName);
            Iterator<Row> rowIterator = sheet.rowIterator();
            if (!rowIterator.hasNext()) {
                throw new RuntimeException("Empty rows");
            }
            return readData(readTitles(rowIterator.next()), rowIterator, evaluator);
        } catch (IOException e) {
            log.error("", e);
            throw new RuntimeException(e);
        }
    }

    private static List<Map<String, String>> readData(List<String> titles, Iterator<Row> rowIterator, FormulaEvaluator evaluator) {
        return Streams.of(rowIterator).map(row -> convertDataToMap(titles, row, evaluator)).collect(Collectors.toList());
    }

    private static Map<String, String> convertDataToMap(List<String> titles, Row row, FormulaEvaluator evaluator) {
        return Streams.of(row.cellIterator())
            .map(cell -> parseCell(cell, titles, evaluator)).collect(Collectors.toMap(ImmutablePair::getLeft, ImmutablePair::getRight, (k1, k2) -> k1));
    }

    private static ImmutablePair<String, String> parseCell(Cell cell, List<String> titles, FormulaEvaluator evaluator) {
        return getSimpleValuePairWithTitle(cell, titles.get(cell.getColumnIndex()), evaluator);
    }

    private static ImmutablePair<String, String> getSimpleValuePairWithTitle(Cell cell,String title, FormulaEvaluator evaluator) {
        CellType cellType = cell.getCellType();

        ImmutablePair<String, String> pair = null;
        switch (cellType) {
            case _NONE:
            case ERROR:
            case BLANK:
                pair = ImmutablePair.of(title, "");
                break;
            case STRING:
                pair = ImmutablePair.of(title, cell.getStringCellValue());
                break;
            case NUMERIC:
                pair = HSSFDateUtil.isCellDateFormatted(cell) ? ImmutablePair.of(title, parseDate(cell)) : ImmutablePair.of(title, parseNumeric(cell.getNumericCellValue()));
                break;
            case BOOLEAN:
                pair = ImmutablePair.of(title, String.valueOf(cell.getBooleanCellValue()));
                break;
            case FORMULA:
                CellValue cellValue = evaluator.evaluate(cell);
                pair = getSimpleValuePairWithTitle(cellValue, cellValue.getCellType(), title);
                break;
        }
        return pair;
    }

    private static ImmutablePair<String, String> getSimpleValuePairWithTitle(CellValue cell, CellType cellType, String title) {
        ImmutablePair<String, String> pair = null;
        switch (cellType) {
            case _NONE:
            case ERROR:
            case BLANK:
                pair = ImmutablePair.of(title, "");
                break;
            case STRING:
                pair = ImmutablePair.of(title, cell.getStringValue());
                break;
            case NUMERIC:
                pair = ImmutablePair.of(title, parseNumeric(cell.getNumberValue()));
                break;
            case BOOLEAN:
                pair = ImmutablePair.of(title, String.valueOf(cell.getBooleanValue()));
                break;
        }
        return pair;
    }

    private static SimpleDateFormat DATE_FORMAT = new SimpleDateFormat("dd/MM/yyyy");

    private static String parseDate(Cell cell) {
        Date dateCellValue = cell.getDateCellValue();
        return DATE_FORMAT.format(dateCellValue);
    }

    private static String parseNumeric(double value) { 
        return String.valueOf(value / Math.round(value) == 1 ? new Double(value).longValue() : value);
    }

    private static List<String> readTitles(Row row) {
        return Streams.of(row.cellIterator()).map(Cell::getStringCellValue).collect(Collectors.toList());
    }
}
