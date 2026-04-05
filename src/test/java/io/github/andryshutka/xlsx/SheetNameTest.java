package io.github.andryshutka.xlsx;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

class SheetNameTest {

    @Test
    void adjustments_hasCorrectLabels() {
        assertEquals("report.adjustment.glovo.adjustments", SheetName.ADJUSTMENTS.getGlovoLabel());
        assertEquals("report.adjustment.partner.adjustments", SheetName.ADJUSTMENTS.getPartnerLabel());
    }

    @Test
    void adjustmentsAfterPeriod_hasCorrectLabels() {
        assertEquals("report.adjustment.glovo.adjustmentAfterPeriod", SheetName.ADJUSTMENTS_AFTER_PERIOD.getGlovoLabel());
        assertEquals("report.adjustment.partner.adjustmentAfterPeriod", SheetName.ADJUSTMENTS_AFTER_PERIOD.getPartnerLabel());
    }

    @Test
    void wtAdjustments_hasCorrectLabels() {
        assertEquals("report.adjustment.glovo.wtAdjustments", SheetName.WT_ADJUSTMENTS.getGlovoLabel());
        assertEquals("report.adjustment.partner.wtAdjustments", SheetName.WT_ADJUSTMENTS.getPartnerLabel());
    }

    @Test
    void wtAdjustmentsAfterPeriod_hasCorrectLabels() {
        assertEquals("report.adjustment.glovo.wtAdjustmentsAfterPeriod", SheetName.WT_ADJUSTMENTS_AFTER_PERIOD.getGlovoLabel());
        assertEquals("report.adjustment.partner.wtAdjustmentsAfterPeriod", SheetName.WT_ADJUSTMENTS_AFTER_PERIOD.getPartnerLabel());
    }

    @Test
    void laasOrders_hasCorrectLabels() {
        assertEquals("report.laasOrders.glovo", SheetName.LAAS_ORDERS.getGlovoLabel());
        assertEquals("report.laasOrders.partner", SheetName.LAAS_ORDERS.getPartnerLabel());
    }

    @Test
    void allEnumValues_arePresent() {
        SheetName[] values = SheetName.values();
        assertEquals(5, values.length);
    }

    @Test
    void valueOf_returnsCorrectEnum() {
        assertEquals(SheetName.ADJUSTMENTS, SheetName.valueOf("ADJUSTMENTS"));
        assertEquals(SheetName.LAAS_ORDERS, SheetName.valueOf("LAAS_ORDERS"));
    }
}
