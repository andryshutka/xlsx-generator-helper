package ui.github.andryshutka.xlsx;

import ui.github.andryshutka.xlsx.annotation.Alignment;
import ui.github.andryshutka.xlsx.annotation.DoNotPrint;
import ui.github.andryshutka.xlsx.annotation.Font;
import ui.github.andryshutka.xlsx.annotation.Formula;
import ui.github.andryshutka.xlsx.annotation.Header;
import ui.github.andryshutka.xlsx.annotation.HyperLink;
import ui.github.andryshutka.xlsx.annotation.Percentage;
import ui.github.andryshutka.xlsx.style.XlsxCellStyleHelper;
import ui.github.andryshutka.xlsx.style.XlsxRowStyleHelper;
import org.apache.commons.compress.archivers.zip.Zip64Mode;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.stream.IntStream;

import static ui.github.andryshutka.xlsx.XlsxConstants.ARIAL_FONT_NAME;
import static ui.github.andryshutka.xlsx.XlsxConstants.DEFAULT_FONT_HEIGHT;
import static org.apache.poi.ss.usermodel.IndexedColors.GREY_25_PERCENT;

/**
 * Component includes methods for convenient work with XLSX files
 */
public class XlsxGeneratorHelper {

  private DataFormat dataFormat;
  private short defaultFormat;
  private short percentageFormat;

  /**
   * creates workbook
   *
   * @return {@link Workbook}
   */
  public Workbook createWorkbook() {
    SXSSFWorkbook workbook = new SXSSFWorkbook();
    workbook.setCompressTempFiles(false);
    workbook.setZip64Mode(Zip64Mode.AsNeeded);
    dataFormat = workbook.createDataFormat();
    defaultFormat = dataFormat.getFormat("### ### ### ##0.00;(### ### ### ##0.00)");
    percentageFormat = HSSFDataFormat.getBuiltinFormat("0.00%");
    return workbook;
  }

  public void disposeWorkbook(Workbook workbook) {
    if (workbook instanceof SXSSFWorkbook) {
      ((SXSSFWorkbook) workbook).dispose();
    }
  }

  public Workbook createWorkbook(int windowSize) {
    SXSSFWorkbook workbook = new SXSSFWorkbook(windowSize);
    workbook.setCompressTempFiles(false);
    workbook.setZip64Mode(Zip64Mode.AsNeeded);
    dataFormat = workbook.createDataFormat();
    defaultFormat = dataFormat.getFormat("### ### ### ##0.00;(### ### ### ##0.00)");
    percentageFormat = HSSFDataFormat.getBuiltinFormat("0.00%");
    return workbook;
  }

  /**
   * Gets or creates sheet for workbook by its name.
   * if there are no sheet with this name in this workbook - new sheet will be created
   *
   * @param workbook  {@link Workbook}
   * @param sheetName {@link String} sheet name
   * @return {@link Sheet} sheet by it's name or new sheet with this name
   */
  public Sheet getOrCreateSheet(Workbook workbook, String sheetName) {
    Sheet sheet = workbook.getSheet(sheetName);
    if (Objects.isNull(sheet)) {
      sheet = workbook.createSheet(sheetName);
    }
    return sheet;
  }

  /**
   * Writes static test into passed sheet to x,y coordinate
   *
   * @param sheet  {@link Sheet} sheet
   * @param rowPos x coordinate - row started from 0
   * @param colPos y coordinate - column started from 0
   * @param text   text to write
   */
  public void writeStaticText(Sheet sheet, int rowPos, int colPos, String text) {
    Workbook workbook = sheet.getWorkbook();
    Cell cell = getOrCreateCell(sheet, rowPos, colPos);
    CellStyle cellStyle = cell.getCellStyle();
    org.apache.poi.ss.usermodel.Font font = workbook.createFont();
    font.setFontHeightInPoints(DEFAULT_FONT_HEIGHT);
    font.setFontName(ARIAL_FONT_NAME);
    cellStyle.setFont(font);
    cellStyle.setWrapText(true);
    cell.setCellStyle(cellStyle);
    cell.setCellValue(text);
  }

  public void appendOrWriteNumber(Sheet sheet, int rowPos, int colPos, BigDecimal number) {
    Cell cell = getOrCreateCell(sheet, rowPos, colPos);
    String existingValue = cell.getCellType() == CellType.STRING ? cell.getStringCellValue() : "";

    CellStyle style = sheet.getWorkbook().createCellStyle();
    style.setAlignment(HorizontalAlignment.RIGHT);

    String numberStr = number.stripTrailingZeros().toPlainString();
    String updatedValue = existingValue.isBlank() ? numberStr : existingValue + " " + numberStr;

    cell.setCellValue(updatedValue);
    cell.setCellStyle(style);
  }

  public List<ReportHeader> getHeadersFromData(XlsxDataInput input) {
    return (input.getHeaders() != null && !input.getHeaders().isEmpty())
        ? input.getHeaders()
        : extractHeaderInfoFromDto(input.getValues());
  }

  public void writeListOfData(XlsxDataInput dataInput) {
    if (dataInput.getValues() == null || dataInput.getValues().isEmpty()) {
      return;
    }
    renderSummary(dataInput);
    Row startRow = getOrCreateRow(dataInput.getSheet(), dataInput.getRowPos());
    List<ReportHeader> headers = getHeadersFromData(dataInput);
    int headersSize = headers.size();
    XlsxRowStyleHelper styleHelper = new XlsxRowStyleHelper(startRow);
    styleHelper
        .withBackgroundColor(GREY_25_PERCENT.getIndex(), headersSize)
        .withHorizontalCenteredText()
        .withVerticalCenteredText()
        .withHeight((short) dataInput.getHeaderHeight())
        .withFont(dataInput.getHeaderFont(), headersSize)
        .withLineBreaks(dataInput.isLineBreaks());
    for (int idx = 0; idx < headersSize; idx++) {
      Cell cell = getOrCreateCell(startRow, idx);
      ReportHeader currentHeader = headers.get(idx);
      writeStaticText(cell, currentHeader.getLabel(), dataInput.isWrapText());
      setColorForSingleColumnIfColorPresent(cell, currentHeader);
    }
    int valuesSize = dataInput.getValues() == null ? 0 : dataInput.getValues().size();

    Map<String, XlsxCellStyleHelper> styles = new HashMap<>();
    if (valuesSize > 0) {
      adjustCellStyle(dataInput, styles);
    }
    if (dataInput.getRowPos() + valuesSize > SpreadsheetVersion.EXCEL2007.getMaxRows()) {
      throw new IllegalArgumentException("Excel rows limit will be reached");
    }
    fillInAllValues(valuesSize, dataInput, styles);
    // this is done last for performance reasons
    adjustSheet(dataInput.getSheet(), headers, startRow);
  }

  private void renderSummary(XlsxDataInput input) {
    int row = input.getRowPos() - 1;
    if (row >= 0 && input.getValues() != null && !input.getValues().isEmpty()) {
      Field[] declaredFields = input.getValues().get(0).getClass().getDeclaredFields();

      List<Field> fields = Arrays.stream(declaredFields)
          .filter(field -> field.isAnnotationPresent(Header.class))
          .toList();

      for (int i = 0; i < fields.size(); i++) {
        Field field = fields.get(i);
        if (field.getAnnotation(Header.class).renderSummary()) {
          Row formulaRow = getOrCreateRow(input.getSheet(), row);
          Cell cell = getOrCreateCell(formulaRow, i);
          styleRow(input.getSheet(), formulaRow);
          int dataStart = row + 3;
          String columnAlpha = getColumnAlpha(i);
          new XlsxCellStyleHelper(cell).withDefaultFont();
          String valueToAdd = (input.getSummaryAppends() != null && !input.getSummaryAppends().isEmpty())
              ? Optional.ofNullable(input.getSummaryAppends().get(columnAlpha)).map(val -> " " + val).orElse("")
              : "";
          String format = String.format("SUM(%s%d:%s%d)" + valueToAdd, columnAlpha, dataStart, columnAlpha, dataStart + input.getValues().size() - 1);
          writeFormula(cell, format);
        }
      }
    }
  }

  private void styleRow(Sheet sheet, Row row) {
    XlsxRowStyleHelper styleHelper = new XlsxRowStyleHelper(row);
    org.apache.poi.ss.usermodel.Font font = sheet.getWorkbook().createFont();
    font.setFontHeightInPoints(DEFAULT_FONT_HEIGHT);
    font.setFontName(ARIAL_FONT_NAME);
    styleHelper.withFont(font, 200);
  }

  private void setColorForSingleColumnIfColorPresent(Cell cell, ReportHeader currentHeader) {
    Map<String, Object> properties = new HashMap<>();
    if (currentHeader.getColor() != null) {
      properties.put(CellUtil.FILL_FOREGROUND_COLOR, currentHeader.getColor().getIndex());
    }
    if (!properties.isEmpty()) {
      CellUtil.setCellStyleProperties(cell, properties);
    }
  }

  /**
   * Writes {@link Double} value(s) in simple row with label ahead. The result will be shown as follows:
   * <pre>
   *     Label | Value1 | Value2 | Value3 | ... |
   * </pre>
   *
   * @param sheet  {@link Sheet} where values should be placed
   * @param rowPos x coordinate - row started from 0
   * @param colPos y coordinate - column started from 0. In this place Label will be printed.
   *               Next values will be printed on (colPos + 1, colPos + 2, ....) coordinate
   * @param label  Ahead label
   * @param values values to print
   */
  public void writeValuesInRow(Sheet sheet, int rowPos, int colPos, String label, Double... values) {
    Cell cell = getOrCreateCell(sheet, rowPos, colPos);
    writeValuesInRow(cell, label, values);
  }

  /**
   * Writes {@link String} value(s) in simple row with label ahead. The result will be shown as follows:
   * <pre>
   *     Label | Value1 | Value2 | Value3 | ... |
   * </pre>
   *
   * @param sheet  {@link Sheet} where values should be placed
   * @param rowPos x coordinate - row started from 0
   * @param colPos y coordinate - column started from 0. In this place Label will be printed.
   *               Next values will be printed on (colPos + 1, colPos + 2, ....) coordinate
   * @param label  Ahead label
   * @param values values to print
   */
  public void writeValuesInRow(Sheet sheet, int rowPos, int colPos, String label, String... values) {
    Cell cell = getOrCreateCell(sheet, rowPos, colPos);
    writeValuesInRow(cell, label, values);
  }

  /**
   * Writes {@link String} value in simple column with label above. The result will be shown as follows:
   * <pre>
   *     _Label_
   *      Value
   * </pre>
   *
   * @param sheet  {@link Sheet} where values should be placed
   * @param rowPos x coordinate - row started from 0
   * @param colPos y coordinate - column started from 0. In this place Label will be printed.
   *               Next value will be printed on (rowPos + 1) coordinate
   * @param label  Above label
   * @param value  value to print
   */
  public void writeValueInColumn(Sheet sheet, int rowPos, int colPos, String label, String value) {
    Cell cell = getOrCreateCell(sheet, rowPos, colPos);
    writeValueInColumn(cell, label, value);
  }

  /**
   * Writes {@link Double} value in simple column with label above. The result will be shown as follows:
   * <pre>
   *     _Label_
   *      Value
   * </pre>
   *
   * @param sheet  {@link Sheet} where values should be placed
   * @param rowPos x coordinate - row started from 0
   * @param colPos y coordinate - column started from 0. In this place Label will be printed.
   *               Next value will be printed on (rowPos + 1) coordinate
   * @param label  Above label
   * @param value  value to print
   */
  public void writeValueInColumn(Sheet sheet, int rowPos, int colPos, String label, Double value) {
    Cell cell = getOrCreateCell(sheet, rowPos, colPos);
    writeValueInColumn(cell, label, value);
  }

  /**
   * Gets or creates row for sheet by its number.
   * if there are no row at this position in this workbook - new row will be created
   *
   * @param sheet  {@link Workbook}
   * @param rowPos {@link Integer} row position started from 0
   * @return {@link Row}
   */
  public Row getOrCreateRow(Sheet sheet, int rowPos) {
    Row rowAtPosition = sheet.getRow(rowPos);
    if (rowAtPosition == null) {
      rowAtPosition = sheet.createRow(rowPos);
    }
    return rowAtPosition;
  }

  /**
   * Gets or creates cell for row by its number.
   * if there are no cell at this position in this workbook - new cell will be created
   *
   * @param row     {@link Row}
   * @param cellPos {@link Integer} cell position started from 0
   * @return {@link Cell}
   */
  public Cell getOrCreateCell(Row row, int cellPos) {
    Cell cellAtPosition = row.getCell(cellPos);
    if (cellAtPosition == null) {
      cellAtPosition = row.createCell(cellPos);
    }
    return cellAtPosition;
  }

  /**
   * Gets or creates cell in (x, y) position. If there are no row and/or cell - new instances will be created
   * if there are no cell at this position in this workbook - new cell will be created
   *
   * @param rowPos  {@link Integer} row position
   * @param cellPos {@link Integer} cell position started from 0
   * @return {@link Cell}
   */
  public Cell getOrCreateCell(Sheet sheet, int rowPos, int cellPos) {
    Row row = getOrCreateRow(sheet, rowPos);
    return getOrCreateCell(row, cellPos);
  }

  /**
   * Adjusts columns size for sheet
   *
   * @param sheet {@link Sheet} to adjust
   */
  public void adjustSheet(Sheet sheet) {
    for (short i = sheet.getRow(getRowWithMaxColumns(sheet)).getFirstCellNum(),
         end = sheet.getRow(getRowWithMaxColumns(sheet)).getLastCellNum(); i < end; i++) {
      sheet.autoSizeColumn(i);
    }
  }

  /**
   * Writes static text into cell
   *
   * @param position cell to write
   * @param text     text to write
   * @param wrapText
   */
  public void writeStaticText(Cell position, String text, boolean wrapText) {
    CellStyle cellStyle = position.getCellStyle();
    cellStyle.setWrapText(wrapText);
    position.setCellValue(text);
  }

  /**
   * Writes static formula into cell
   *
   * @param position cell to write
   * @param formula  to write
   */
  public void writeFormula(Cell position, String formula) {
    position.setCellFormula(formula);
    CellStyle cellStyle = position.getCellStyle();
    cellStyle.setDataFormat(defaultFormat);
    org.apache.poi.ss.usermodel.Font font = position.getSheet().getWorkbook().createFont();
    font.setFontHeightInPoints(DEFAULT_FONT_HEIGHT);
    font.setFontName(ARIAL_FONT_NAME);
    cellStyle.setFont(font);
    position.setCellStyle(cellStyle);
  }

  private void writeHyperLink(Cell cell, String value) {
    CellStyle cellStyle = cell.getCellStyle();
    cellStyle.setDataFormat(defaultFormat);
    org.apache.poi.ss.usermodel.Font font = cell.getSheet().getWorkbook().createFont();
    font.setFontHeightInPoints(DEFAULT_FONT_HEIGHT);
    font.setFontName(ARIAL_FONT_NAME);
    cellStyle.setFont(font);

    cell.setCellValue(value);
    CreationHelper creationHelper = cell.getSheet().getWorkbook().getCreationHelper();
    Hyperlink link = creationHelper.createHyperlink(HyperlinkType.URL);

    link.setAddress(value);
    cell.setHyperlink(link);
    cell.setCellStyle(cellStyle);
  }

  /**
   * Return column 'letter(s)' by cell
   *
   * @param cell
   * @return 'A', 'B', etc
   */
  public String getColumnAlpha(Cell cell) {
    return CellReference.convertNumToColString(cell.getColumnIndex());
  }

  /**
   * Return column 'letter(s)' by column index
   *
   * @param columnIndex
   * @return 'A', 'B', etc
   */
  public String getColumnAlpha(int columnIndex) {
    return CellReference.convertNumToColString(columnIndex);
  }

  public void mergeCell(Sheet sheet, Cell cell, int verticalCount, int horizontalCount) {
    int rowIndex = cell.getRowIndex();
    int columnIndex = cell.getColumnIndex();
    sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex + verticalCount, columnIndex, columnIndex + horizontalCount));
  }

  public void writeValueInColumn(Sheet sheet, int rowPos, int colPos, Double value) {
    Cell cell = getOrCreateCell(sheet, rowPos, colPos);
    writeNumber(cell, value);
  }

  /**
   * Writes {@link Date} value into cell. Format is MM/dd/yyyy
   *
   * @param position cell
   * @param date     date value
   */
  private void writeDate(Cell position, Date date) {
    position.setCellValue(date);
    Workbook currentWb = position.getSheet().getWorkbook();
    CellStyle cellStyle = position.getCellStyle();
    CreationHelper createHelper = currentWb.getCreationHelper();
    cellStyle.setDataFormat(
        createHelper.createDataFormat().getFormat(XlsxConstants.DATE_FORMAT));
    position.setCellStyle(cellStyle);
  }

  private void writeDate(Cell position, LocalDate date) {
    if (date != null) {
      String dateFormatted = date.format(XlsxConstants.DATE_FORMATTER);
      writeStaticText(position, dateFormatted, false);
    }
  }

  private void writeDate(Cell position, LocalDateTime date) {
    if (date != null) {
      String dateFormatted = date.format(XlsxConstants.DATE_TIME_FORMATTER);
      writeStaticText(position, dateFormatted, false);
    }
  }

  /**
   * Writes {@link Long} value into cell
   *
   * @param position cell
   * @param value    number
   */
  private void writeNumber(Cell position, Long value) {
    position.setCellValue(value);
  }

  /**
   * Writes {@link Double} value into cell
   *
   * @param position cell
   * @param value    number
   */
  private void writeNumber(Cell position, Double value) {
    position.setCellValue(value);
    CellStyle cellStyle = position.getCellStyle();
    cellStyle.setDataFormat(defaultFormat);
    position.setCellStyle(cellStyle);
  }

  /**
   * Writes {@link Double} value into cell with percentage formatting
   *
   * @param position cell
   * @param value    number
   */
  private void writePercentage(Cell position, Double value) {
    position.setCellValue(value);
    CellStyle cellStyle = position.getCellStyle();
    cellStyle.setDataFormat(percentageFormat);
    position.setCellStyle(cellStyle);
  }

  /**
   * Writes {@link Double} value(s) in simple row with label ahead. The result will be shown as follows:
   * <pre>
   *     Label | Value1 | Value2 | Value3 | ... |
   * </pre>
   *
   * @param position {@link Cell} cell
   * @param text     Ahead label
   * @param values   values to print
   */
  private void writeValuesInRow(Cell position, String text, Double... values) {
    writeStaticText(position, text, true);
    int valuesSize = values.length;
    int i = 0;
    while (i < valuesSize) {
      Cell nextCell = getOrCreateCell(position.getRow(), position.getColumnIndex() + (i + 1));
      writeNumber(nextCell, values[i]);
      i++;
    }
  }

  private void writeSingleValue(Cell cell, Object fieldValue, Field field, boolean wrapText) {
    if (fieldValue instanceof Number) {
      if (field.isAnnotationPresent(Percentage.class)) {
        writePercentage(cell, ((Number) fieldValue).doubleValue());
      } else {
        if (fieldValue instanceof BigDecimal || fieldValue instanceof Double || fieldValue instanceof Float) {
          writeNumber(cell, ((Number) fieldValue).doubleValue());
        } else if (fieldValue instanceof Long || fieldValue instanceof Integer) {
          writeNumber(cell, ((Number) fieldValue).longValue());
        }
      }
    } else if (fieldValue instanceof String) {
      if (field.isAnnotationPresent(Formula.class)) {
        writeFormula(cell, (String) fieldValue);
      } else if (field.isAnnotationPresent(HyperLink.class)) {
        writeHyperLink(cell, (String) fieldValue);
      } else {
        writeStaticText(cell, (String) fieldValue, wrapText);
      }
    } else if (fieldValue instanceof Boolean) {
      cell.setCellValue((Boolean) fieldValue);
    } else if (fieldValue instanceof LocalDate) {
      writeDate(cell, (LocalDate) fieldValue);
    } else if (fieldValue instanceof LocalDateTime) {
      writeDate(cell, (LocalDateTime) fieldValue);
    } else {
      writeStaticText(cell, Objects.isNull(fieldValue) ? "" : fieldValue.toString(), wrapText);
    }
  }

  /**
   * Writes {@link String} value(s) in simple row with label ahead. The result will be shown as follows:
   * <pre>
   *     Label | Value1 | Value2 | Value3 | ... |
   * </pre>
   *
   * @param position {@link Cell} cell
   * @param text     Ahead label
   * @param values   values to print
   */
  private void writeValuesInRow(Cell position, String text, String... values) {
    writeStaticText(position, text, true);
    int valuesSize = values.length;
    int i = 0;
    while (i < valuesSize) {
      Cell nextCell = getOrCreateCell(position.getRow(), position.getColumnIndex() + (i + 1));
      writeStaticText(nextCell, values[i], true);
      i++;
    }
  }

  /**
   * Writes {@link String} value in simple column with label above. The result will be shown as follows:
   * <pre>
   *     _Label_
   *      Value
   * </pre>
   *
   * @param position {@link Cell} cell
   * @param text     Above label
   * @param value    value to print
   */
  private void writeValueInColumn(Cell position, String text, String value) {
    writeStaticText(position, text, true);
    Row lowerRow = getOrCreateRow(position.getSheet(), position.getRowIndex() + 1);
    Cell lowerCell = getOrCreateCell(lowerRow, position.getColumnIndex());
    writeStaticText(lowerCell, value, true);
  }

  /**
   * Writes {@link Double} value in simple column with label above. The result will be shown as follows:
   * <pre>
   *     _Label_
   *      Value
   * </pre>
   *
   * @param position {@link Cell} cell
   * @param text     Above label
   * @param value    value to print
   */
  private void writeValueInColumn(Cell position, String text, Double value) {
    writeStaticText(position, text, true);
    Row lowerRow = getOrCreateRow(position.getSheet(), position.getRowIndex() + 1);
    Cell lowerCell = getOrCreateCell(lowerRow, position.getColumnIndex());
    writeNumber(lowerCell, value);
  }

  /**
   * gets row with maximum amount of created cells in
   *
   * @param sheet sheet to investigate
   * @return row number
   */
  private int getRowWithMaxColumns(Sheet sheet) {
    int colNumber = 0;
    int rowNumber = 0;
    Iterator<Row> rowIterator = sheet.iterator();
    while (rowIterator.hasNext()) {
      Row row = rowIterator.next();
      short lastCellNum = row.getLastCellNum();
      if (lastCellNum > colNumber) {
        colNumber = lastCellNum;
        rowNumber = row.getRowNum();
      }
    }
    return rowNumber;
  }

  private void adjustSheet(Sheet sheet, List<ReportHeader> headers, Row startRow) {
    int headersSize = headers.size();
    for (int idx = 0; idx < headersSize; idx++) {
      Cell cell = getOrCreateCell(startRow, idx);
      ReportHeader currentHeader = headers.get(idx);
      sheet.setColumnWidth(cell.getColumnIndex(), 255 * currentHeader.getWidth());
    }
  }

  private void adjustCellStyle(XlsxDataInput dataInput, Map<String, XlsxCellStyleHelper> styles) {
    Field[] dFields = dataInput.getValues().get(0).getClass().getDeclaredFields();
    IntStream.range(0, dFields.length).forEach(fieldIdx -> {
      Field field = dFields[fieldIdx];
      XlsxCellStyleHelper cellStyle = styles.get(field.getName());
      if (cellStyle == null) {
        cellStyle = new XlsxCellStyleHelper(dataInput.getSheet().getWorkbook());
      }
      cellStyle.withAlignment(field.getDeclaredAnnotation(Alignment.class));
      cellStyle.withFont(field.getDeclaredAnnotation(Font.class));
      cellStyle.withLineBreaks(dataInput.isLineBreaks());
      styles.put(field.getName(), cellStyle);
    });
  }

  private void fillInAllValues(int valuesSize, XlsxDataInput dataInput, Map<String, XlsxCellStyleHelper> styles) {
    int fieldsSize = 0;
    boolean wrapText = dataInput.isWrapText();
    List<?> values = dataInput.getValues();
    List<Field> declaredFields = new ArrayList<>();
    Map<Field, Boolean> isFieldAllowedByHeaderMap = new HashMap<>();
    Map<Field, Boolean> isDoNotPrintMap = new HashMap<>();
    if (valuesSize > 0) {
      Object value = values.get(0);
      declaredFields = getFields(value);
      fieldsSize = declaredFields.size();
      for (int fieldIdx = 0; fieldIdx < fieldsSize; fieldIdx++) {
        Field field = declaredFields.get(fieldIdx);
        field.setAccessible(true);
        isFieldAllowedByHeaderMap.put(field, isFieldAllowedByHeader(field));
        isDoNotPrintMap.put(field, field.isAnnotationPresent(DoNotPrint.class));
      }
    }
    for (int valIndex = 0; valIndex < valuesSize; valIndex++) {
      Object value = values.get(valIndex);
      int fieldSkip = 0;
      for (int fieldIdx = 0; fieldIdx < fieldsSize; fieldIdx++) {
        Field field = declaredFields.get(fieldIdx);
        if (isFieldAllowedByHeaderMap.get(field) && !isDoNotPrintMap.get(field)) {
          Object fieldValue = getFieldValue(field, value);
          Cell cell = getOrCreateCell(dataInput.getSheet(), dataInput.getRowPos() + valIndex + 1, fieldIdx - fieldSkip);
          cell.setCellStyle(styles.get(field.getName()).getCellStyle());
          writeSingleValue(cell, fieldValue, field, wrapText);
        } else {
          fieldSkip++;
        }
      }
    }
  }

  private Object getFieldValue(Field field, Object target) {
    try {
      field.setAccessible(true);
      return field.get(target);
    } catch (IllegalAccessException e) {
      throw new RuntimeException("Failed to access field: " + field.getName(), e);
    }
  }

  private boolean isFieldAllowedByHeader(Field field) {
    if (field.isAnnotationPresent(Header.class)) {
      return true;
    }
    return true;
  }

  private List<ReportHeader> extractHeaderInfoFromDto(List<?> values) {
    Class<?> valueType = values.get(0).getClass();
    Field[] declaredFields = valueType.getDeclaredFields();
    List<ReportHeader> headers = new LinkedList<>();
    for (int i = 0; i < declaredFields.length; i++) {
      Field declaredField = declaredFields[i];
      Header header = declaredField.getDeclaredAnnotation(Header.class);
      if (Objects.nonNull(header)) {
        headers.add(buildReportHeader(header));
      } else {
        if (!declaredField.isAnnotationPresent(DoNotPrint.class)) {
          headers.add(ReportHeader.of("missing header for field: " + declaredField.getName(), 60));
        }
      }
    }
    return headers;
  }

  private ReportHeader buildReportHeader(Header header) {
    String title = header.label();
    int width = header.widthSize();
    return ReportHeader.builder()
        .label(title)
        .width(width < 1
            ? Arrays.stream(title.split(" ")).mapToInt(String::length).max().orElse(10) + 3
            : width)
        .build();
  }

  /**
   * Do not use DTO hierarchy!
   */
  private <T> List<Field> getFields(T t) {
    List<Field> fields = new ArrayList<>();
    Class clazz = t.getClass();
    fields.addAll(0, Arrays.asList(clazz.getDeclaredFields()));
    return fields;
  }
}
