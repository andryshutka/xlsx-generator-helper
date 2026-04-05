package io.github.andryshutka.xlsx;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;

class XlsxDataInputTest {

    private Workbook workbook;
    private Sheet sheet;

    @BeforeEach
    void setUp() {
        workbook = new SXSSFWorkbook();
        sheet = workbook.createSheet("Test");
    }

    @AfterEach
    void tearDown() throws IOException {
        workbook.close();
    }

    @Test
    void builder_setsAllFields() {
        List<ReportHeader> headers = List.of(ReportHeader.of("Col1", 10));
        List<String> values = List.of("a", "b");
        Map<String, String> summaryAppends = Map.of("A", "+10");

        XlsxDataInput input = XlsxDataInput.builder()
                .sheet(sheet)
                .rowPos(3)
                .headers(headers)
                .values(values)
                .lineBreaks(true)
                .headerHeight(30)
                .wrapText(true)
                .summaryAppends(summaryAppends)
                .build();

        assertSame(sheet, input.getSheet());
        assertEquals(3, input.getRowPos());
        assertEquals(headers, input.getHeaders());
        assertEquals(values, input.getValues());
        assertTrue(input.isLineBreaks());
        assertEquals(30, input.getHeaderHeight());
        assertTrue(input.isWrapText());
        assertEquals(summaryAppends, input.getSummaryAppends());
    }

    @Test
    void builder_defaults_areFalseAndZero() {
        XlsxDataInput input = XlsxDataInput.builder()
                .sheet(sheet)
                .build();

        assertEquals(0, input.getRowPos());
        assertEquals(0, input.getHeaderHeight());
        assertFalse(input.isLineBreaks());
        assertFalse(input.isWrapText());
        assertNull(input.getHeaders());
        assertNull(input.getValues());
        assertNull(input.getHeaderColor());
        assertNull(input.getHeaderFont());
        assertNull(input.getSummaryAppends());
    }

    @Test
    void builder_withEmptyHeaders_returnsEmptyList() {
        XlsxDataInput input = XlsxDataInput.builder()
                .sheet(sheet)
                .headers(List.of())
                .build();

        assertTrue(input.getHeaders().isEmpty());
    }
}
