package io.github.andryshutka.xlsx;

import org.junit.jupiter.api.Test;

import java.time.LocalDateTime;

import static org.junit.jupiter.api.Assertions.*;

class XlsxConstantsTest {

    @Test
    void fontNames_areCorrect() {
        assertEquals("Calibri", XlsxConstants.DEFAULT_FONT_NAME);
        assertEquals("Arial", XlsxConstants.ARIAL_FONT_NAME);
    }

    @Test
    void fontHeights_areCorrect() {
        assertEquals((short) 8, XlsxConstants.DEFAULT_FONT_HEIGHT);
        assertEquals((short) 8, XlsxConstants.DEFAULT_HEADER_FONT_HEIGHT);
    }

    @Test
    void dateFormat_isCorrect() {
        assertEquals("dd.MM.yyyy", XlsxConstants.DATE_FORMAT);
    }

    @Test
    void dateTimeFormat_isCorrect() {
        assertEquals("dd.MM.yyyy HH:mm", XlsxConstants.DATE_TIME_FORMAT);
    }

    @Test
    void dateTimeFormatter_formatsCorrectly() {
        LocalDateTime dateTime = LocalDateTime.of(2024, 3, 15, 9, 5);
        String formatted = XlsxConstants.DATE_TIME_FORMATTER.format(dateTime);
        assertEquals("15.03.2024 09:05", formatted);
    }

    @Test
    void dateTimeFormatter_isConsistentWithDateTimeFormat() {
        LocalDateTime dateTime = LocalDateTime.of(2023, 12, 1, 23, 59);
        String formatted = XlsxConstants.DATE_TIME_FORMATTER.format(dateTime);
        assertEquals("01.12.2023 23:59", formatted);
    }
}
