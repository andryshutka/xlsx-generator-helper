package ui.github.andryshutka.xlsx;

import org.apache.poi.ss.usermodel.IndexedColors;

public class ReportHeader {

  private final String label;
  private final int width;
  private final IndexedColors color;

  public static ReportHeader of(String label, int width) {
    return new ReportHeader(label, width, null);
  }

  public static ReportHeader of(String label, int width, IndexedColors color) {
    return new ReportHeader(label, width, color);
  }

  private ReportHeader(String label, int width, IndexedColors color) {
    this.label = label;
    this.width = width;
    this.color = color;
  }

  public String getLabel() {
    return label;
  }

  public int getWidth() {
    return width;
  }

  public IndexedColors getColor() {
    return color;
  }

  public static Builder builder() {
    return new Builder();
  }

  public static class Builder {
    private String label;
    private int width;
    private IndexedColors color;

    public Builder label(String label) {
      this.label = label;
      return this;
    }

    public Builder width(int width) {
      this.width = width;
      return this;
    }

    public Builder color(IndexedColors color) {
      this.color = color;
      return this;
    }

    public ReportHeader build() {
      return new ReportHeader(label, width, color);
    }
  }
}
