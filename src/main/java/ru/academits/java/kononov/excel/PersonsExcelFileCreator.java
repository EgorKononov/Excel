package ru.academits.java.kononov.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import ru.academits.java.kononov.excel.person.Person;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class PersonsExcelFileCreator {
    private static final int COLUMN_INDEX_NAME = 0;
    private static final int COLUMN_INDEX_SURNAME = 1;
    private static final int COLUMN_INDEX_AGE = 2;
    private static final int COLUMN_INDEX_PHONE = 3;

    private static Sheet sheet;

    private static Row header;
    private static CellStyle headerCellStyle;

    private static Row row;
    private static CellStyle style;

    private PersonsExcelFileCreator() {
    }

    static void create(List<Person> persons, String filePath) throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            createHeader(workbook);
            addPersonsRows(persons, workbook);
            writeToFile(filePath, workbook);
        }
    }

    private static void createHeader(Workbook workbook) {
        sheet = workbook.createSheet("Persons");
        header = sheet.createRow(0);

        createHeaderStyle(workbook);

        addHeaderCell(COLUMN_INDEX_NAME, "Name");
        addHeaderCell(COLUMN_INDEX_SURNAME, "Surname");
        addHeaderCell(COLUMN_INDEX_AGE, "Age");
        addHeaderCell(COLUMN_INDEX_PHONE, "Phone");
    }

    private static void createHeaderStyle(Workbook workbook) {
        headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        headerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        XSSFFont font = createHeaderFont((XSSFWorkbook) workbook);

        headerCellStyle.setFont(font);
    }

    private static XSSFFont createHeaderFont(XSSFWorkbook workbook) {
        XSSFFont font = workbook.createFont();
        font.setFontHeightInPoints((short) 12);
        font.setBold(true);

        return font;
    }

    private static void addHeaderCell(int i, String value) {
        Cell headerCell = header.createCell(i);
        headerCell.setCellValue(value);
        headerCell.setCellStyle(headerCellStyle);
    }

    private static void addPersonsRows(List<Person> persons, Workbook workbook) {
        style = workbook.createCellStyle();
        style.setWrapText(true);

        int i = 1;

        for (Person person : persons) {
            row = sheet.createRow(i);

            addCell(COLUMN_INDEX_NAME, person.getName());
            addCell(COLUMN_INDEX_SURNAME, person.getSurname());
            addCell(COLUMN_INDEX_AGE, person.getAge());
            addCell(COLUMN_INDEX_PHONE, person.getPhone());

            ++i;
        }

        setColumnSize();
    }

    private static void addCell(int i, String value) {
        Cell cell = createCell(i);
        cell.setCellValue(value);
    }

    private static void addCell(int i, double value) {
        Cell cell = createCell(i);
        cell.setCellValue(value);
    }

    private static Cell createCell(int i) {
        Cell cell = row.createCell(i);
        cell.setCellStyle(style);

        return cell;
    }

    private static void setColumnSize() {
        for (int j = 0; j < 4; j++) {
            sheet.autoSizeColumn(j);
        }
    }

    private static void writeToFile(String filePath, Workbook workbook) throws IOException {
        FileOutputStream outputStream = new FileOutputStream(filePath);
        workbook.write(outputStream);
    }
}
