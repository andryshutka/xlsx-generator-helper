package io.github.andryshutka.xlsx.style;


import org.apache.poi.ss.usermodel.BorderExtent;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.PropertyTemplate;

/**
 * Helper class for styling xlsx cell ranges
 */
public class XlsxCellRangeStyleHelper {

  /**
   * Add borders for a cell range
   * @param sheet {@link Sheet}
   * @param firstRow Index of first row
   * @param lastRow Index of last row (inclusive)
   * @param firstCol Index of first column
   * @param lastCol Index of last column (inclusive)
   */
  public void withBorder(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol) {
    PropertyTemplate propertyTemplate = new PropertyTemplate();
    propertyTemplate.drawBorders(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol), BorderStyle.MEDIUM, BorderExtent.OUTSIDE);
    propertyTemplate.applyBorders(sheet);
  }

}
