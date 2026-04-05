package io.github.andryshutka.xlsx;

import io.github.andryshutka.xlsx.annotation.Header;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

class XlsxGeneratorHelperTest {

    private XlsxGeneratorHelper helper;
    private Workbook workbook;
    private Sheet sheet;

    @BeforeEach
    void setUp() {
        helper = new XlsxGeneratorHelper();
        workbook = helper.createWorkbook();
        sheet = workbook.createSheet("TestSheet");
    }

    @AfterEach
    void tearDown() {
        helper.disposeWorkbook(workbook);
    }

    // ── Workbook creation ──────────────────────────────────────────────────────

    @Test
    void createWorkbook_returnsSXSSFWorkbook() {
        assertInstanceOf(SXSSFWorkbook.class, workbook);
    }

    @Test
    void createWorkbook_withWindowSize_returnsSXSSFWorkbook() {
        Workbook wb = helper.createWorkbook(100);
        assertInstanceOf(SXSSFWorkbook.class, wb);
        helper.disposeWorkbook(wb);
    }

    @Test
    void disposeWorkbook_onNonSXSSF_doesNotThrow() {
        // disposeWorkbook should be a no-op for non-SXSSF workbooks
        assertDoesNotThrow(() -> helper.disposeWorkbook(null));
    }

    // ── Sheet get/create ───────────────────────────────────────────────────────

    @Test
    void getOrCreateSheet_returnsExistingSheet() {
        Sheet existing = workbook.createSheet("Existing");
        Sheet result = helper.getOrCreateSheet(workbook, "Existing");
        assertSame(existing, result);
    }

    @Test
    void getOrCreateSheet_createsNewSheetWhenAbsent() {
        Sheet result = helper.getOrCreateSheet(workbook, "NewSheet");
        assertNotNull(result);
        assertEquals("NewSheet", result.getSheetName());
        assertSame(result, workbook.getSheet("NewSheet"));
    }

    @Test
    void getOrCreateSheet_calledTwice_returnsSameSheet() {
        Sheet first = helper.getOrCreateSheet(workbook, "Repeated");
        Sheet second = helper.getOrCreateSheet(workbook, "Repeated");
        assertSame(first, second);
    }

    // ── Column alpha conversion ────────────────────────────────────────────────

    @Test
    void getColumnAlpha_byIndex_returnsCorrectLetters() {
        assertEquals("A", helper.getColumnAlpha(0));
        assertEquals("B", helper.getColumnAlpha(1));
        assertEquals("Z", helper.getColumnAlpha(25));
        assertEquals("AA", helper.getColumnAlpha(26));
        assertEquals("AZ", helper.getColumnAlpha(51));
    }

    @Test
    void getColumnAlpha_byCell_returnsCorrectLetter() {
        Row row = sheet.createRow(0);
        Cell cellA = row.createCell(0);
        Cell cellC = row.createCell(2);
        assertEquals("A", helper.getColumnAlpha(cellA));
        assertEquals("C", helper.getColumnAlpha(cellC));
    }

    // ── Static text writing ────────────────────────────────────────────────────

    @Test
    void writeStaticText_writesValueToCell() {
        helper.writeStaticText(sheet, 0, 0, "Hello");
        Cell cell = sheet.getRow(0).getCell(0);
        assertNotNull(cell);
        assertEquals("Hello", cell.getStringCellValue());
    }

    @Test
    void writeStaticText_createsRowAndCellIfAbsent() {
        assertNull(sheet.getRow(5));
        helper.writeStaticText(sheet, 5, 3, "Test");
        assertNotNull(sheet.getRow(5));
        assertEquals("Test", sheet.getRow(5).getCell(3).getStringCellValue());
    }

    @Test
    void writeStaticText_overwritesExistingValue() {
        helper.writeStaticText(sheet, 0, 0, "First");
        helper.writeStaticText(sheet, 0, 0, "Second");
        assertEquals("Second", sheet.getRow(0).getCell(0).getStringCellValue());
    }

    // ── appendOrWriteNumber ────────────────────────────────────────────────────

    @Test
    void appendOrWriteNumber_writesNumberToEmptyCell() {
        helper.appendOrWriteNumber(sheet, 0, 0, new BigDecimal("42.5"));
        Cell cell = sheet.getRow(0).getCell(0);
        assertEquals("42.5", cell.getStringCellValue());
    }

    @Test
    void appendOrWriteNumber_appendsToExistingStringValue() {
        helper.writeStaticText(sheet, 0, 0, "Existing");
        helper.appendOrWriteNumber(sheet, 0, 0, new BigDecimal("10"));
        assertEquals("Existing 10", sheet.getRow(0).getCell(0).getStringCellValue());
    }

    @Test
    void appendOrWriteNumber_stripsTrailingZeros() {
        helper.appendOrWriteNumber(sheet, 0, 0, new BigDecimal("5.00"));
        assertEquals("5", sheet.getRow(0).getCell(0).getStringCellValue());
    }

    // ── writeValuesInRow ───────────────────────────────────────────────────────

    @Test
    void writeValuesInRow_withDoubles_writesLabelAndValues() {
        helper.writeValuesInRow(sheet, 0, 0, "Revenue", 100.0, 200.0, 300.0);
        Row row = sheet.getRow(0);
        assertEquals("Revenue", row.getCell(0).getStringCellValue());
        assertEquals(100.0, row.getCell(1).getNumericCellValue(), 0.001);
        assertEquals(200.0, row.getCell(2).getNumericCellValue(), 0.001);
        assertEquals(300.0, row.getCell(3).getNumericCellValue(), 0.001);
    }

    @Test
    void writeValuesInRow_withStrings_writesLabelAndValues() {
        helper.writeValuesInRow(sheet, 1, 0, "Header", "A", "B", "C");
        Row row = sheet.getRow(1);
        assertEquals("Header", row.getCell(0).getStringCellValue());
        assertEquals("A", row.getCell(1).getStringCellValue());
        assertEquals("B", row.getCell(2).getStringCellValue());
        assertEquals("C", row.getCell(3).getStringCellValue());
    }

    // ── writeValueInColumn ─────────────────────────────────────────────────────

    @Test
    void writeValueInColumn_withString_writesLabelAboveAndValueBelow() {
        helper.writeValueInColumn(sheet, 0, 0, "Label", "Value");
        assertEquals("Label", sheet.getRow(0).getCell(0).getStringCellValue());
        assertEquals("Value", sheet.getRow(1).getCell(0).getStringCellValue());
    }

    @Test
    void writeValueInColumn_withDouble_writesLabelAboveAndNumberBelow() {
        helper.writeValueInColumn(sheet, 0, 0, "Total", 99.9);
        assertEquals("Total", sheet.getRow(0).getCell(0).getStringCellValue());
        assertEquals(99.9, sheet.getRow(1).getCell(0).getNumericCellValue(), 0.001);
    }

    @Test
    void writeValueInColumn_withDoubleDirectly_writesNumber() {
        helper.writeValueInColumn(sheet, 2, 1, 55.5);
        assertEquals(55.5, sheet.getRow(2).getCell(1).getNumericCellValue(), 0.001);
    }

    // ── mergeCell ─────────────────────────────────────────────────────────────

    @Test
    void mergeCell_addsMergedRegion() {
        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);
        helper.mergeCell(sheet, cell, 1, 2);
        assertEquals(1, sheet.getNumMergedRegions());
    }

    // ── getHeadersFromData ─────────────────────────────────────────────────────

    @Test
    void getHeadersFromData_withExplicitHeaders_returnsThoseHeaders() {
        List<ReportHeader> headers = List.of(ReportHeader.of("Col", 10));
        XlsxDataInput input = XlsxDataInput.builder()
                .sheet(sheet)
                .headers(headers)
                .values(List.of())
                .build();
        assertEquals(headers, helper.getHeadersFromData(input));
    }

    @Test
    void getHeadersFromData_withNullHeaders_extractsFromAnnotations() {
        List<SampleDto> values = List.of(new SampleDto("Alice", 30));
        XlsxDataInput input = XlsxDataInput.builder()
                .sheet(sheet)
                .headers(null)
                .values(values)
                .build();
        List<ReportHeader> headers = helper.getHeadersFromData(input);
        assertEquals(2, headers.size());
        assertEquals("Name", headers.get(0).getLabel());
        assertEquals("Age", headers.get(1).getLabel());
    }

    // ── writeListOfData ────────────────────────────────────────────────────────

    @Test
    void writeListOfData_withEmptyValues_doesNotWriteAnything() {
        XlsxDataInput input = XlsxDataInput.builder()
                .sheet(sheet)
                .rowPos(0)
                .values(List.of())
                .build();
        helper.writeListOfData(input);
        assertNull(sheet.getRow(0));
    }

    @Test
    void writeListOfData_writesHeaderRowAndDataRows() {
        List<SampleDto> values = List.of(
                new SampleDto("Alice", 30),
                new SampleDto("Bob", 25)
        );
        XlsxDataInput input = XlsxDataInput.builder()
                .sheet(sheet)
                .rowPos(0)
                .values(values)
                .headerHeight(20)
                .build();
        helper.writeListOfData(input);

        // Header row at rowPos=0
        Row headerRow = sheet.getRow(0);
        assertNotNull(headerRow);
        assertEquals("Name", headerRow.getCell(0).getStringCellValue());
        assertEquals("Age", headerRow.getCell(1).getStringCellValue());

        // Data rows at rowPos+1 and rowPos+2
        Row dataRow1 = sheet.getRow(1);
        assertNotNull(dataRow1);
        assertEquals("Alice", dataRow1.getCell(0).getStringCellValue());
        assertEquals(30.0, dataRow1.getCell(1).getNumericCellValue(), 0.001);

        Row dataRow2 = sheet.getRow(2);
        assertNotNull(dataRow2);
        assertEquals("Bob", dataRow2.getCell(0).getStringCellValue());
        assertEquals(25.0, dataRow2.getCell(1).getNumericCellValue(), 0.001);
    }

    @Test
    void writeListOfData_withLocalDateField_writesFormattedDate() {
        List<DtoWithDate> values = List.of(new DtoWithDate(LocalDate.of(2024, 6, 15)));
        XlsxDataInput input = XlsxDataInput.builder()
                .sheet(sheet)
                .rowPos(0)
                .values(values)
                .headerHeight(20)
                .build();
        helper.writeListOfData(input);

        Row dataRow = sheet.getRow(1);
        assertNotNull(dataRow);
        assertEquals("15.06.2024", dataRow.getCell(0).getStringCellValue());
    }

    @Test
    void writeListOfData_withLocalDateTimeField_writesFormattedDateTime() {
        List<DtoWithDateTime> values = List.of(new DtoWithDateTime(LocalDateTime.of(2024, 6, 15, 10, 30)));
        XlsxDataInput input = XlsxDataInput.builder()
                .sheet(sheet)
                .rowPos(0)
                .values(values)
                .headerHeight(20)
                .build();
        helper.writeListOfData(input);

        Row dataRow = sheet.getRow(1);
        assertNotNull(dataRow);
        assertEquals("15.06.2024 10:30", dataRow.getCell(0).getStringCellValue());
    }

    @Test
    void writeListOfData_exceedingExcelRowLimit_throwsIllegalArgumentException() {
        // rowPos near the limit so that rowPos + values.size() > max rows
        int maxRows = 1048576; // EXCEL2007 max
        List<SampleDto> values = List.of(new SampleDto("X", 1));
        XlsxDataInput input = XlsxDataInput.builder()
                .sheet(sheet)
                .rowPos(maxRows)
                .values(values)
                .headerHeight(20)
                .build();
        assertThrows(IllegalArgumentException.class, () -> helper.writeListOfData(input));
    }

    // ── Sample DTOs ────────────────────────────────────────────────────────────

    static class SampleDto {
        @Header(label = "Name")
        private final String name;

        @Header(label = "Age")
        private final int age;

        SampleDto(String name, int age) {
            this.name = name;
            this.age = age;
        }
    }

    static class DtoWithDate {
        @Header(label = "Date")
        private final LocalDate date;

        DtoWithDate(LocalDate date) {
            this.date = date;
        }
    }

    static class DtoWithDateTime {
        @Header(label = "DateTime")
        private final LocalDateTime dateTime;

        DtoWithDateTime(LocalDateTime dateTime) {
            this.dateTime = dateTime;
        }
    }
}
