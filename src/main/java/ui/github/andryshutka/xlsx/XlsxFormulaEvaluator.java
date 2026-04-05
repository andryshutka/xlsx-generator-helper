package ui.github.andryshutka.xlsx;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.util.stream.IntStream;

public class XlsxFormulaEvaluator {

  private final SXSSFWorkbook workbook;

  public XlsxFormulaEvaluator(SXSSFWorkbook workbook) {
    this.workbook = workbook;
  }

  public void evaluate() {
    int numberOfSheets = workbook.getNumberOfSheets();
    IntStream.range(0, numberOfSheets).forEach(i -> {
      Sheet sheet = workbook.getSheetAt(i);
      sheet.setForceFormulaRecalculation(true);
    });
  }
}
