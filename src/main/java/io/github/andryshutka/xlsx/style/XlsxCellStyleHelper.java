package io.github.andryshutka.xlsx.style;

import io.github.andryshutka.xlsx.XlsxConstants;
import io.github.andryshutka.xlsx.annotation.Alignment;
import io.github.andryshutka.xlsx.annotation.Font;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellUtil;

/**
 * Helper class for styling xlsx cells
 */
public class XlsxCellStyleHelper {

  /**
   * cell to style
   */
  private Cell cell;
  private final Workbook workbook;
  private final CellStyle cellStyle;

  /**
   * Creates styler wrapper for cell
   * @param cell {@link Cell}
   */
  public XlsxCellStyleHelper(Cell cell) {
    this.cell = cell;
    this.workbook = this.cell.getSheet().getWorkbook();
    this.cellStyle = workbook.createCellStyle();
  }

  public XlsxCellStyleHelper(Workbook workbook) {
    this.workbook = workbook;
    this.cellStyle = workbook.createCellStyle();
  }

  public CellStyle getCellStyle() {
    return cellStyle;
  }

  /**
   * sets background color for cell
   * @param background color code  @see {@link org.apache.poi.ss.usermodel.IndexedColors}
   * @return this
   */
  public XlsxCellStyleHelper withBackgroundColor(short background) {
    CellUtil.setCellStyleProperty(cell, CellUtil.FILL_FOREGROUND_COLOR, background);
    CellUtil.setCellStyleProperty(cell, CellUtil.FILL_PATTERN, FillPatternType.SOLID_FOREGROUND);
    return this;
  }

  public XlsxCellStyleHelper withFont(Cell cell, Font fontStyle) {
    org.apache.poi.ss.usermodel.Font font = workbook.createFont();
    font.setBold(fontStyle.bold());
    font.setItalic(fontStyle.italic());
    font.setFontName(fontStyle.value());
    font.setFontHeightInPoints(fontStyle.size());
    CellUtil.setFont(cell, font);
    return this;
  }

  /**
   * aligns text
   */
  public void withAlignment(Alignment alignment) {
    if (alignment != null) {
      cellStyle.setAlignment(alignment.value());
    }
  }

  public void withFont(Font fontStyle) {
    if (fontStyle != null) {
      org.apache.poi.ss.usermodel.Font font = workbook.createFont();
      font.setBold(fontStyle.bold());
      font.setItalic(fontStyle.italic());
      font.setFontName(fontStyle.value());
      font.setFontHeightInPoints(fontStyle.size());
      cellStyle.setFont(font);
    }
  }

  public void withDefaultFont() {
    org.apache.poi.ss.usermodel.Font font = workbook.createFont();
    font.setBold(false);
    font.setItalic(false);
    font.setFontName(XlsxConstants.ARIAL_FONT_NAME);
    font.setFontHeightInPoints((short) 8);
    cellStyle.setFont(font);
    cell.setCellStyle(cellStyle);
  }

  /**
   * allows to display a line break
   */
  public void withLineBreaks(boolean withLineBreaks) {
    cellStyle.setWrapText(withLineBreaks);
  }
}
