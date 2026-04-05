# xlsx-generator

A lightweight Java library for generating XLSX files using Apache POI

## Requirements

- Java 17+
- Maven 3.8+

## Installation

Add the following dependency to your `pom.xml`:

```xml
<dependency>
    <groupId>io.github.andryshutka</groupId>
    <artifactId>xlsx-generator</artifactId>
    <version>1.0.0</version>
</dependency>
```

## Quick Start

```java
import io.github.andryshutka.xlsx.XlsxGeneratorHelper;
import io.github.andryshutka.xlsx.XlsxDataInput;
import io.github.andryshutka.xlsx.annotation.Header;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.util.List;

// 1. Define your data class
public class OrderRow {

    @Header(label = "Order ID")
    private String orderId;

    @Header(label = "Amount")
    private BigDecimal amount;

    @Header(label = "Date")
    private LocalDate date;
}

// 2. Generate the workbook
XlsxGeneratorHelper helper = new XlsxGeneratorHelper();
Workbook workbook = helper.createWorkbook();

XlsxDataInput input = XlsxDataInput.builder()
        .sheet(helper.getOrCreateSheet(workbook, "Orders"))
        .values(List.of(new OrderRow("ORD-1", new BigDecimal("99.99"), LocalDate.now())))
        .startRowPosition(0)
        .startColPosition(0)
        .build();

helper.writeListOfData(input);

try (FileOutputStream out = new FileOutputStream("orders.xlsx")) {
    workbook.write(out);
}
helper.disposeWorkbook(workbook);
```

## Annotations

Annotations are placed on fields of your data DTO to control how each column is rendered.

### `@Header`

Marks a field as a column and configures its header cell.

| Attribute                   | Type            | Default           | Description                                    |
|-----------------------------|-----------------|-------------------|------------------------------------------------|
| `label`                     | `String`        | *(required)*      | Column header text                             |
| `widthSize`                 | `int`           | `-1` (ignored)    | Fixed column width in POI units                |
| `widthAsHeaderLength`       | `boolean`       | `true`            | Auto-size column to header label length        |
| `widthAsAverageValueLength` | `boolean`       | `false`           | Auto-size column based on average value length |
| `color`                     | `IndexedColors` | `GREY_25_PERCENT` | Header cell background color                   |
| `renderSummary`             | `boolean`       | `false`           | Include this column in summary row             |
| `forCountry`                | `String[]`      | `{}`              | Render column only for specified country codes |

```java
@Header(label = "Revenue", color = IndexedColors.LIGHT_BLUE, widthSize = 20)
private BigDecimal revenue;
```

### `@Font`

Customizes the font of a cell's value.

| Attribute | Type      | Default   | Description         |
|-----------|-----------|-----------|---------------------|
| `value`   | `String`  | `"Arial"` | Font family name    |
| `bold`    | `boolean` | `false`   | Bold text           |
| `italic`  | `boolean` | `false`   | Italic text         |
| `size`    | `short`   | `8`       | Font size in points |

```java
@Header(label = "Product")
@Font(bold = true, size = 10)
private String productName;
```

### `@Alignment`

Sets the horizontal alignment of a cell.

```java
@Header(label = "SKU")
@Alignment(HorizontalAlignment.CENTER)
private String sku;
```

### `@Percentage`

Formats the numeric value as a percentage (`0.00%`).

```java
@Header(label = "Discount")
@Percentage
private BigDecimal discount;
```

### `@Formula`

Marks a `String` field whose value is a raw Excel formula (e.g. `"SUM(A1:A10)"`).

```java
@Header(label = "Total")
@Formula
private String totalFormula;
```

### `@HyperLink`

Renders the cell value as a clickable hyperlink.

```java
@Header(label = "URL")
@HyperLink
private String url;
```

### `@DoNotPrint`

Excludes a field from the generated output entirely.

```java
@DoNotPrint
private String internalId;
```

## Supported Field Types

The following Java types are handled automatically when writing data rows:

| Java Type             | Excel output                      |
|-----------------------|-----------------------------------|
| `String`              | Text cell                         |
| `Integer` / `int`     | Numeric cell                      |
| `Long` / `long`       | Numeric cell                      |
| `Double` / `double`   | Numeric cell                      |
| `BigDecimal`          | Numeric cell (formatted)          |
| `LocalDate`           | Text cell (`dd.MM.yyyy`)          |
| `LocalDateTime`       | Text cell (`dd.MM.yyyy HH:mm:ss`) |
| `Date`                | Text cell (`dd.MM.yyyy HH:mm:ss`) |
| `Boolean` / `boolean` | Text cell (`true` / `false`)      |

## `XlsxDataInput` Builder

`XlsxDataInput` is the main configuration object passed to `writeListOfData`.

```java
XlsxDataInput input = XlsxDataInput.builder()
        .sheet(sheet)             // target Sheet (required)
        .values(rows)             // List<?> of data objects (required)
        .headers(customHeaders)   // List<ReportHeader> — overrides @Header annotations
        .startRowPosition(2)      // first data row index (0-based), default 0
        .startColPosition(1)      // first data column index (0-based), default 0
        .rowLimit(1_000_000)      // max rows before splitting (default: XLSX limit)
        .build();
```

## `ReportHeader` — Manual Headers

Instead of annotations you can supply headers programmatically:

```java
ReportHeader header = ReportHeader.builder()
        .label("Order ID")
        .width(15)
        .color(IndexedColors.LIGHT_YELLOW.getIndex())
        .build();
```

## Other Utilities

### `XlsxGeneratorHelper`

| Method                                              | Description                                              |
|-----------------------------------------------------|----------------------------------------------------------|
| `createWorkbook()`                                  | Creates a streaming `SXSSFWorkbook`                      |
| `createWorkbook(int windowSize)`                    | Creates a workbook with a custom in-memory row window    |
| `disposeWorkbook(Workbook)`                         | Releases temp files used by `SXSSFWorkbook`              |
| `getOrCreateSheet(Workbook, String)`                | Gets an existing sheet or creates a new one              |
| `writeStaticText(Sheet, row, col, text)`            | Writes a plain text cell at the given position           |
| `writeListOfData(XlsxDataInput)`                    | Writes a list of annotated DTOs as rows                  |
| `getHeadersFromData(XlsxDataInput)`                 | Extracts headers from annotations or the input's list    |
| `mergeRegion(Sheet, firstRow, lastRow, firstCol, lastCol)` | Merges a cell range                             |

### `XlsxFormulaEvaluator`

Recalculates all formulas in a workbook after writing data:

```java
XlsxFormulaEvaluator evaluator = new XlsxFormulaEvaluator();
evaluator.evaluate(workbook);
```

## License

[Apache License, Version 2.0](https://www.apache.org/licenses/LICENSE-2.0)
