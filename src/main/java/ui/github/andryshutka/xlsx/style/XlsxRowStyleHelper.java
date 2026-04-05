package ui.github.andryshutka.xlsx.style;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.XSSFColor;

import java.util.Objects;

import static org.apache.poi.ss.util.CellUtil.BORDER_BOTTOM;
import static org.apache.poi.ss.util.CellUtil.BORDER_TOP;
import static org.apache.poi.ss.util.CellUtil.FILL_FOREGROUND_COLOR;
import static org.apache.poi.ss.util.CellUtil.FILL_PATTERN;
import static org.apache.poi.ss.util.CellUtil.FONT;
import static org.apache.poi.ss.util.CellUtil.WRAP_TEXT;

/**
 * Helper class for styling xlsx rows
 */
public class XlsxRowStyleHelper {

  /**
   * row to style
   */
  private final Row row;
  private final Workbook workbook;


  /**
   * Creates styler wrapper for row
   * @param row {@link Row}
   */
  public XlsxRowStyleHelper(Row row) {
    this.row = row;
    this.workbook = this.row.getSheet().getWorkbook();
  }

  public XlsxRowStyleHelper withFont(Font font, int len) {
    for (short i = 0; i < len; i++) {
      Cell cell = getCell(i);
      CellStyle cellStyle = cell.getCellStyle();
      cellStyle.setFont(font);
    }
    return this;
  }

  /**
   * aligns text to center horizontally
   * @return this
   */
  public XlsxRowStyleHelper withHorizontalCenteredText() {
    if (!row.isFormatted()) {
      row.setRowStyle(row.getSheet().getWorkbook().createCellStyle());
    }
    row.getRowStyle().setAlignment(HorizontalAlignment.CENTER);
    return this;
  }

  /**
   * aligns text to center vertically
   * @return this
   */
  public XlsxRowStyleHelper withVerticalCenteredText() {
    if (!row.isFormatted()) {
      row.setRowStyle(row.getSheet().getWorkbook().createCellStyle());
    }
    row.getRowStyle().setVerticalAlignment(VerticalAlignment.CENTER);
    return this;
  }

  /**
   * aligns text to center vertically
   * @return this
   */
  public XlsxRowStyleHelper withHeight() {
    row.setHeight((short)-1);
    return this;
  }


  /**
   * sets background color for row
   * @param background color code  @see {@link org.apache.poi.ss.usermodel.IndexedColors}
   * @return this
   */
  public XlsxRowStyleHelper withBackgroundColor(short background) {
    apply(FILL_FOREGROUND_COLOR, background);
    apply(FILL_PATTERN, FillPatternType.SOLID_FOREGROUND);
    return this;
  }

  public XlsxRowStyleHelper withBackgroundColor(XSSFColor color, int len) {
    for (short i = 0; i < len; i++) {
      Cell cell = getCell(i);
      CellUtil.setCellStyleProperty(cell, FILL_FOREGROUND_COLOR, color);
      CellUtil.setCellStyleProperty(cell, FILL_PATTERN, FillPatternType.SOLID_FOREGROUND);
    }
    return this;
  }

  public XlsxRowStyleHelper withBackgroundColor(short color, int len) {
    for (short i = 0; i < len; i++) {
      Cell cell = getCell(i);
      CellUtil.setCellStyleProperty(cell, FILL_FOREGROUND_COLOR, color);
      CellUtil.setCellStyleProperty(cell, FILL_PATTERN, FillPatternType.SOLID_FOREGROUND);
    }
    return this;
  }


  /**
   * sets background color for whole row
   * @param background color code  @see {@link org.apache.poi.ss.usermodel.IndexedColors}
   * @return this
   */
  public XlsxRowStyleHelper withBackgroundColorWholeRow(short background) {
    for (short i = 0,
         end = 100; i < end; i++) {
      Cell cell = getCell(i);
      CellUtil.setCellStyleProperty(cell, FILL_FOREGROUND_COLOR, background);
      CellUtil.setCellStyleProperty(cell, FILL_PATTERN, FillPatternType.SOLID_FOREGROUND);
    }
    return this;
  }

  /**
   * sets height for row
   * @param height pixels
   * @return this
   */
  public XlsxRowStyleHelper withHeight(short height) {
    row.setHeight(height);
    return this;
  }

  /**
   * sets top border for created cells in row
   * @return this
   */
  public XlsxRowStyleHelper withTopBorder() {
    apply(BORDER_TOP, BorderStyle.MEDIUM);
    return this;
  }

  /**
   * sets bottom border for created cells in row
   * @return this
   */
  public XlsxRowStyleHelper withBottomBorder() {
    apply(BORDER_BOTTOM, BorderStyle.MEDIUM);
    return this;
  }

  private Cell getCell(int cellNum) {
    Cell cellAtPosition = row.getCell(cellNum);
    if (Objects.isNull(cellAtPosition)) {
      cellAtPosition = row.createCell(cellNum);
    }
    return cellAtPosition;
  }

  private void apply(String property, Object value) {
    for (short i = row.getFirstCellNum(),
         end = row.getLastCellNum(); i < end; i++) {
      CellUtil.setCellStyleProperty(getCell(i), property, value);
    }
  }

  public void withFont(String fontName) {
    Font font = workbook.createFont();
    font.setFontName(fontName);
    apply(FONT, font);
  }

  /**
   * allows to display a line break
   */
  public void withLineBreaks(boolean lineBreaks) {
    apply(WRAP_TEXT, lineBreaks);
  }
}
