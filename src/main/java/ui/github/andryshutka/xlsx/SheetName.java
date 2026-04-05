package ui.github.andryshutka.xlsx;

public enum SheetName {

    ADJUSTMENTS(
        "report.adjustment.glovo.adjustments",
        "report.adjustment.partner.adjustments"),
    ADJUSTMENTS_AFTER_PERIOD(
        "report.adjustment.glovo.adjustmentAfterPeriod",
        "report.adjustment.partner.adjustmentAfterPeriod"),
    WT_ADJUSTMENTS(
        "report.adjustment.glovo.wtAdjustments",
        "report.adjustment.partner.wtAdjustments"),
    WT_ADJUSTMENTS_AFTER_PERIOD(
        "report.adjustment.glovo.wtAdjustmentsAfterPeriod",
        "report.adjustment.partner.wtAdjustmentsAfterPeriod"),
    LAAS_ORDERS(
        "report.laasOrders.glovo",
        "report.laasOrders.partner");

    private final String glovoLabel;
    private final String partnerLabel;

    SheetName(String glovoLabel, String partnerLabel) {
        this.glovoLabel = glovoLabel;
        this.partnerLabel = partnerLabel;
    }

    public String getGlovoLabel() {
        return glovoLabel;
    }

    public String getPartnerLabel() {
        return partnerLabel;
    }

}
