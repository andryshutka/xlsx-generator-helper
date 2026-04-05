package io.github.andryshutka.xlsx;

import java.time.format.DateTimeFormatter;

public class XlsxConstants {

  public static final String DEFAULT_FONT_NAME = "Calibri";
  public static final String ARIAL_FONT_NAME = "Arial";
  public static final Short DEFAULT_FONT_HEIGHT = 8;
  public static final Short DEFAULT_HEADER_FONT_HEIGHT = 8;
  public static final String DATE_FORMAT = "dd.MM.yyyy";

  public static final String DATE_TIME_FORMAT = "dd.MM.yyyy HH:mm";
  public static final DateTimeFormatter DATE_FORMATTER = DateTimeFormatter.ofPattern(DATE_FORMAT);
  public static final DateTimeFormatter DATE_TIME_FORMATTER = DateTimeFormatter.ofPattern(DATE_TIME_FORMAT);

  private XlsxConstants() {
  }

}
