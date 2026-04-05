package ui.github.andryshutka.xlsx;

import org.apache.poi.ss.usermodel.IndexedColors;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

class ReportHeaderTest {

    @Test
    void factoryOf_withLabelAndWidth_setsFieldsAndNullColor() {
        ReportHeader header = ReportHeader.of("Name", 20);

        assertEquals("Name", header.getLabel());
        assertEquals(20, header.getWidth());
        assertNull(header.getColor());
    }

    @Test
    void factoryOf_withLabelWidthAndColor_setsAllFields() {
        ReportHeader header = ReportHeader.of("Amount", 15, IndexedColors.RED);

        assertEquals("Amount", header.getLabel());
        assertEquals(15, header.getWidth());
        assertEquals(IndexedColors.RED, header.getColor());
    }

    @Test
    void builder_setsAllFields() {
        ReportHeader header = ReportHeader.builder()
                .label("Total")
                .width(25)
                .color(IndexedColors.BLUE)
                .build();

        assertEquals("Total", header.getLabel());
        assertEquals(25, header.getWidth());
        assertEquals(IndexedColors.BLUE, header.getColor());
    }

    @Test
    void builder_withoutColor_colorIsNull() {
        ReportHeader header = ReportHeader.builder()
                .label("Label")
                .width(10)
                .build();

        assertNull(header.getColor());
    }

    @Test
    void builder_withZeroWidth_widthIsZero() {
        ReportHeader header = ReportHeader.builder().label("X").width(0).build();
        assertEquals(0, header.getWidth());
    }
}
