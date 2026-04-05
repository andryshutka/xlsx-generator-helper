package io.github.andryshutka.xlsx;

import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.IOException;

import static org.junit.jupiter.api.Assertions.*;

class XlsxFormulaEvaluatorTest {

    @Test
    void evaluate_setsForceFormulaRecalculationOnAllSheets() throws IOException {
        try (SXSSFWorkbook workbook = new SXSSFWorkbook()) {
            workbook.createSheet("Sheet1");
            workbook.createSheet("Sheet2");
            workbook.createSheet("Sheet3");

            new XlsxFormulaEvaluator(workbook).evaluate();

            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                assertTrue(workbook.getSheetAt(i).getForceFormulaRecalculation(),
                        "Sheet " + i + " should have forceFormulaRecalculation=true");
            }
        }
    }

    @Test
    void evaluate_onEmptyWorkbook_doesNotThrow() throws IOException {
        try (SXSSFWorkbook workbook = new SXSSFWorkbook()) {
            assertDoesNotThrow(() -> new XlsxFormulaEvaluator(workbook).evaluate());
        }
    }

    @Test
    void evaluate_onSingleSheet_setsForceFormulaRecalculation() throws IOException {
        try (SXSSFWorkbook workbook = new SXSSFWorkbook()) {
            workbook.createSheet("OnlySheet");

            new XlsxFormulaEvaluator(workbook).evaluate();

            assertTrue(workbook.getSheetAt(0).getForceFormulaRecalculation());
        }
    }
}
