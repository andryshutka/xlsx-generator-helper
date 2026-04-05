package ui.github.andryshutka.xlsx;

import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFColor;

import java.util.List;
import java.util.Map;

public class XlsxDataInput {

  private final Sheet sheet;
  private final int rowPos;
  private final List<ReportHeader> headers;
  private final List<?> values;
  private final XSSFColor headerColor;
  private final Font headerFont;
  private final boolean lineBreaks;
  private final int headerHeight;
  private final boolean wrapText;
  // if field has Header annotation with renderSummary = true
  private final Map<String, String> summaryAppends;

  private XlsxDataInput(Builder builder) {
    this.sheet = builder.sheet;
    this.rowPos = builder.rowPos;
    this.headers = builder.headers;
    this.values = builder.values;
    this.headerColor = builder.headerColor;
    this.headerFont = builder.headerFont;
    this.lineBreaks = builder.lineBreaks;
    this.headerHeight = builder.headerHeight;
    this.wrapText = builder.wrapText;
    this.summaryAppends = builder.summaryAppends;
  }

  public Sheet getSheet() {
    return sheet;
  }

  public int getRowPos() {
    return rowPos;
  }

  public List<ReportHeader> getHeaders() {
    return headers;
  }

  public List<?> getValues() {
    return values;
  }

  public XSSFColor getHeaderColor() {
    return headerColor;
  }

  public Font getHeaderFont() {
    return headerFont;
  }

  public boolean isLineBreaks() {
    return lineBreaks;
  }

  public int getHeaderHeight() {
    return headerHeight;
  }

  public boolean isWrapText() {
    return wrapText;
  }

  public Map<String, String> getSummaryAppends() {
    return summaryAppends;
  }

  public static Builder builder() {
    return new Builder();
  }

  public static class Builder {
    private Sheet sheet;
    private int rowPos;
    private List<ReportHeader> headers;
    private List<?> values;
    private XSSFColor headerColor;
    private Font headerFont;
    private boolean lineBreaks;
    private int headerHeight;
    private boolean wrapText;
    private Map<String, String> summaryAppends;

    public Builder sheet(Sheet sheet) {
      this.sheet = sheet;
      return this;
    }

    public Builder rowPos(int rowPos) {
      this.rowPos = rowPos;
      return this;
    }

    public Builder headers(List<ReportHeader> headers) {
      this.headers = headers;
      return this;
    }

    public Builder values(List<?> values) {
      this.values = values;
      return this;
    }

    public Builder headerColor(XSSFColor headerColor) {
      this.headerColor = headerColor;
      return this;
    }

    public Builder headerFont(Font headerFont) {
      this.headerFont = headerFont;
      return this;
    }

    public Builder lineBreaks(boolean lineBreaks) {
      this.lineBreaks = lineBreaks;
      return this;
    }

    public Builder headerHeight(int headerHeight) {
      this.headerHeight = headerHeight;
      return this;
    }

    public Builder wrapText(boolean wrapText) {
      this.wrapText = wrapText;
      return this;
    }

    public Builder summaryAppends(Map<String, String> summaryAppends) {
      this.summaryAppends = summaryAppends;
      return this;
    }

    public XlsxDataInput build() {
      return new XlsxDataInput(this);
    }
  }
}
