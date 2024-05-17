package aaa.utils.poi;

import static java.lang.Math.max;
import static java.util.Optional.ofNullable;
import static java.util.stream.Collectors.toList;
import static java.util.stream.StreamSupport.stream;

import aaa.nvl.Nvl;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.Date;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Objects;
import java.util.Set;
import java.util.stream.Stream;
import lombok.NonNull;
import lombok.SneakyThrows;
import lombok.experimental.UtilityClass;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HeaderFooter;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.PageMargin;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.SSFExtraTools;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;

/** Apache POI utils */
@UtilityClass
public class PoiUtils {

  public static final int MIN_WIDTH = 256 * 4;

  /** Максимальное значение высоты страницы по-умолчанию */
  public static final double MAX_PAGE_HEIGHT_IN_POINTS = 842;

  public static final double MAX_PAGE_WIDTH_IN_POINTS = 27500;

  /** Get row with certain index, create if non */
  public static Row getOrCreateRow(Sheet sheet, int rowIndex) {
    Row row = sheet.getRow(rowIndex);
    if (row == null) {
      row = sheet.createRow(rowIndex);
    }
    return row;
  }

  /** Get cell with certain index, create if non */
  public static Cell getOrCreateCell(Row row, int cellIndex) {
    Cell cell = row.getCell(cellIndex);
    if (cell == null) {
      cell = row.createCell(cellIndex);
    }
    return cell;
  }

  /** Get cell with certain address, create if non */
  public static Cell getOrCreateCell(Sheet sheet, int rowIndex, short cellIndex) {
    return getOrCreateCell(getOrCreateRow(sheet, rowIndex), cellIndex);
  }

  //
  // createCustomCell functions family
  // Creates cell at certain coordinates with style and value
  //

  public static Cell createCustomCell(Row row, long column, CellStyle style, RichTextString value) {
    Cell cell = getOrCreateCell(row, (short) column);
    if (value != null) {
      cell.setCellValue(value);
    }
    if (style != null) {
      cell.setCellStyle(style);
    }
    return cell;
  }

  public static Cell createCustomCell(Row row, long column, CellStyle style, String value) {
    return createCustomCell(row, column, style, new HSSFRichTextString(value));
  }

  public static Cell createCustomCell(Row row, long column, CellStyle style, Long value) {
    Cell cell = getOrCreateCell(row, (short) column);
    if (value != null) {
      cell.setCellValue(value);
    }
    if (style != null) {
      cell.setCellStyle(style);
    }
    return cell;
  }

  public static Cell createCustomCell(Row row, long column, CellStyle style, Double value) {
    Cell cell = getOrCreateCell(row, (short) column);
    if (value != null) {
      cell.setCellValue(value);
    }
    if (style != null) {
      cell.setCellStyle(style);
    }
    return cell;
  }

  public static Cell createCustomCell(
      Sheet sheet, long row, long column, CellStyle style, RichTextString value) {
    return createCustomCell(getOrCreateRow(sheet, (int) row), column, style, value);
  }

  public static Cell createCustomCell(
      Sheet sheet, long row, long column, CellStyle style, String value) {
    return createCustomCell(getOrCreateRow(sheet, (int) row), column, style, value);
  }

  public static Cell createCustomCell(
      Sheet sheet, long row, long column, CellStyle style, long value) {
    return createCustomCell(getOrCreateRow(sheet, (int) row), column, style, value);
  }

  public static Cell createCustomCell(
      Sheet sheet, long row, long column, CellStyle style, Double value) {
    return createCustomCell(getOrCreateRow(sheet, (int) row), column, style, value);
  }

  public static void clearCellValue(Sheet sheet, long row, long column) {
    getOrCreateCell(sheet, (int) row, (short) column).setCellValue((RichTextString) null);
  }

  /** Flushes workbook to byte array */
  @SneakyThrows
  public static byte[] flushWorkBook(Workbook workBook) {
    ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
    workBook.write(outputStream);
    return outputStream.toByteArray();
  }

  /**
   * Add pages counter to single sheet of workbook
   *
   * @param sheet Sheet to which numbering will be added
   * @param pageWord Word "page" (for internationalization purposes)
   * @param totalPagesDelimiter Delimiter between current page number and total pages count number
   */
  @SuppressWarnings("nls")
  public static void addPagesNumbers(Sheet sheet, String pageWord, String totalPagesDelimiter) {
    sheet
        .getFooter()
        .setRight(
            (StringUtils.isNotBlank(pageWord) ? pageWord + " " : "")
                + HeaderFooter.page()
                + StringUtils.defaultString(totalPagesDelimiter)
                + HeaderFooter.numPages());
  }

  /**
   * Makes named cell in Excel workbook
   *
   * @param workBook Workbook to be modified
   * @param sheet Cell sheet
   * @param name Name of the cell
   * @param rowIndex Row index
   * @param columnIndex Column index
   */
  public static void makeNamedCell(
      Workbook workBook, Sheet sheet, String name, long rowIndex, long columnIndex) {
    makeNamedCell(workBook, workBook.getSheetIndex(sheet), name, rowIndex, columnIndex);
  }

  /**
   * Makes named cell in Excel workbook
   *
   * @param workBook Workbook to be modified
   * @param sheetIndex Cell sheet
   * @param name Name of the cell
   * @param rowIndex Row index
   * @param columnIndex Column index
   */
  public static void makeNamedCell(
      Workbook workBook, int sheetIndex, String name, long rowIndex, long columnIndex) {
    makeNamedCell(workBook, workBook.getSheetName(sheetIndex), name, rowIndex, columnIndex);
  }

  /**
   * Makes named cell in Excel workbook
   *
   * @param workBook Workbook to be modified
   * @param sheetName Cell sheet name
   * @param name Name of the cell
   * @param rowIndex Row index
   * @param columnIndex Column index
   */
  public static void makeNamedCell(
      Workbook workBook, String sheetName, String name, long rowIndex, long columnIndex) {
    Name namedCell = workBook.createName();
    namedCell.setNameName(name);
    namedCell.setRefersToFormula(
        new CellReference(sheetName, (int) rowIndex, (short) columnIndex, false, false)
            .formatAsString());
  }

  /** Get reference to named cell */
  public static CellReference getNamedCellReference(Workbook workbook, String cellName) {
    return ofNullable(workbook.getName(cellName))
        .map(Name::getRefersToFormula)
        .map(CellReference::new)
        .orElse(null);
  }

  /** Get named cell value */
  public static Cell getNamedCell(Workbook workbook, String cellName) {
    CellReference cellReference = getNamedCellReference(workbook, cellName);
    return cellReference == null
        ? null
        : workbook
            .getSheet(cellReference.getSheetName())
            .getRow(cellReference.getRow())
            .getCell(cellReference.getCol());
  }

  public static void setNamedCellValue(
      Workbook workbook, String cellName, RichTextString value, CellStyle style) {
    Cell cell = getNamedCell(workbook, cellName);
    if (cell != null && value != null) {
      cell.setCellValue(value);
      if (style != null) {
        cell.setCellStyle(style);
      }
    }
  }

  public static void setNamedCellValue(
      Workbook workbook, String cellName, String value, CellStyle style) {
    setNamedCellValue(workbook, cellName, new HSSFRichTextString(value), style);
  }

  public static void setNamedCellValue(Workbook workbook, String cellName, String value) {
    setNamedCellValue(workbook, cellName, new HSSFRichTextString(value));
  }

  public static void setNamedCellValue(Workbook workbook, String cellName, RichTextString value) {
    setNamedCellValue(workbook, cellName, value, null);
  }

  public static void setNamedCellValue(
      Workbook workbook, String cellName, Date value, CellStyle style) {
    Cell cell = getNamedCell(workbook, cellName);
    if (cell != null && value != null) {
      cell.setCellValue(value);
      if (style != null) {
        cell.setCellStyle(style);
      }
    }
  }

  public static void setNamedCellValue(Workbook workbook, String cellName, Long value) {
    Cell cell = getNamedCell(workbook, cellName);
    if (cell != null) {
      cell.setCellValue(value);
    }
  }

  public static void setNamedCellValue(Workbook workbook, String cellName, Double value) {
    Cell cell = getNamedCell(workbook, cellName);
    if (cell != null) {
      cell.setCellValue(value);
    }
  }

  public static int getNamedCellRowIndex(Workbook workbook, String cellName) {
    CellReference cellReference = getCellReference(workbook, cellName);
    return cellReference == null ? -1 : cellReference.getRow();
  }

  public static int getNamedCellColumnIndex(Workbook workbook, String cellName) {
    CellReference cellReference = getCellReference(workbook, cellName);
    return cellReference == null ? -1 : cellReference.getCol();
  }

  private static CellReference getCellReference(Workbook workbook, String cellName) {
    return ofNullable(workbook.getName(cellName))
        .map(
            name ->
                new AreaReference(name.getRefersToFormula(), SpreadsheetVersion.EXCEL97)
                    .getFirstCell())
        .orElse(null);
  }

  /** Получение крайнего (максимального) номера столбца на листе */
  public static int getMaxColumnNumber(Sheet sheet) {
    int result = 0;
    for (Row row : sheet) {
      result = Nvl.max(result, (int) row.getLastCellNum());
    }
    return result;
  }

  /** Проверка равенства описаний шрифтов */
  public static boolean fontsEquals(Font font1, Font font2) {
    return (font1.getItalic() == font2.getItalic())
        && (font1.getStrikeout() == font2.getStrikeout())
        && (font1.getBold() == font2.getBold())
        && (font1.getFontHeightInPoints() == font2.getFontHeightInPoints())
        && (font1.getColor() == font2.getColor())
        && (StringUtils.equals(font1.getFontName(), font2.getFontName()))
        && (font1.getTypeOffset() == font2.getTypeOffset())
        && (font1.getUnderline() == font2.getUnderline());
  }

  /** Создание копии описания шрифта в другой рабочей книге */
  public static Font copyFontTo(Font sourceFont, Workbook targetWorkbook) {
    if (sourceFont == null) {
      return null;
    }
    for (short i = 0; i < targetWorkbook.getNumberOfFonts(); ++i) {
      Font font = targetWorkbook.getFontAt(i);
      if (fontsEquals(font, sourceFont)) {
        return font;
      }
    }
    Font targetFont = targetWorkbook.createFont();
    targetFont.setBold(sourceFont.getBold());
    targetFont.setColor(sourceFont.getColor());
    targetFont.setFontHeightInPoints(sourceFont.getFontHeightInPoints());
    targetFont.setFontName(sourceFont.getFontName());
    targetFont.setItalic(sourceFont.getItalic());
    targetFont.setStrikeout(sourceFont.getStrikeout());
    targetFont.setTypeOffset(sourceFont.getTypeOffset());
    targetFont.setUnderline(sourceFont.getUnderline());
    return targetFont;
  }

  static Font getFont(Workbook workbook, CellStyle style) {
    return workbook.getFontAt(style.getFontIndex());
  }

  /** Проверка равенства 2х стилей */
  public static boolean styleEquals(
      Workbook workbook1, CellStyle style1, Workbook workbook2, CellStyle style2) {
    return (style1.getAlignment() == style2.getAlignment())
        && (style1.getBorderBottom() == style2.getBorderBottom())
        && (style1.getBorderTop() == style2.getBorderTop())
        && (style1.getBorderLeft() == style2.getBorderLeft())
        && (style1.getBorderRight() == style2.getBorderRight())
        && (fontsEquals(getFont(workbook1, style1), getFont(workbook2, style2)))
        && (style1.getFillBackgroundColor() == style2.getFillBackgroundColor())
        && (style1.getFillForegroundColor() == style2.getFillForegroundColor())
        && (style1.getIndention() == style2.getIndention())
        && (style1.getRotation() == style2.getRotation())
        && (style1.getVerticalAlignment() == style2.getVerticalAlignment())
        && (style1.getWrapText() == style2.getWrapText());
  }

  @SneakyThrows
  public static Workbook workbookFromBytes(@NonNull byte[] data) {
    return WorkbookFactory.create(new ByteArrayInputStream(data));
  }

  private static class OffsetHolder {
    int rowOffset = 0;
    int columnOffset = 0;
    int numberInPage = 0;
    int numberInRow = 0;
  }

  /** Создание копии стиля оформления в другой рабочей книге */
  public static CellStyle copyStyleTo(
      Workbook sourceWorkbook, CellStyle sourceStyle, Workbook targetWorkbook) {
    if (sourceStyle == null) {
      return null;
    }
    for (short i = 0; i < targetWorkbook.getNumCellStyles(); ++i) {
      CellStyle style = targetWorkbook.getCellStyleAt(i);
      if (styleEquals(sourceWorkbook, sourceStyle, targetWorkbook, style)) {
        return style;
      }
    }

    CellStyle targetStyle = targetWorkbook.createCellStyle();
    targetStyle.setAlignment(sourceStyle.getAlignment());
    targetStyle.setBorderBottom(sourceStyle.getBorderBottom());
    targetStyle.setBorderTop(sourceStyle.getBorderTop());
    targetStyle.setBorderLeft(sourceStyle.getBorderLeft());
    targetStyle.setBorderRight(sourceStyle.getBorderRight());
    Font targetFont = copyFontTo(getFont(sourceWorkbook, sourceStyle), targetWorkbook);
    if (targetFont != null) {
      targetStyle.setFont(targetFont);
    }
    targetStyle.setFillBackgroundColor(sourceStyle.getFillBackgroundColor());
    targetStyle.setFillForegroundColor(sourceStyle.getFillForegroundColor());
    targetStyle.setIndention(sourceStyle.getIndention());
    targetStyle.setRotation(sourceStyle.getRotation());
    targetStyle.setVerticalAlignment(sourceStyle.getVerticalAlignment());
    targetStyle.setWrapText(sourceStyle.getWrapText());
    return targetStyle;
  }

  public static int getNextUncreatedRowIndex(Sheet sheet, boolean first) {
    return first ? 0 : sheet.getLastRowNum() + 1;
  }

  private static class ColumnSizingData {
    public Set<Integer> columns = new HashSet<>();
    public int maxColumnIndex = -1;
  }

  /**
   * Добавление данных из некоторой книги Excel на заданный рабочий лист Excel. <br>
   * Добавление происходит "в столбик". <br>
   * Подразумевается, что данные для в исходной рабочей книге сформированы по одному шаблону
   *
   * @param sourceWorkbook Исходная рабочая книга
   * @param targetSheet Лист на рабочей книге, на котором необходимо свести данные
   * @param targetWorkbook Целевая рабочая книга
   */
  private static boolean collectToSingle(
      Workbook sourceWorkbook,
      Sheet targetSheet,
      Workbook targetWorkbook,
      boolean forcePageBreaks,
      int maxPerPage,
      int maxPerRow,
      Double maxPageHeightInPoints,
      Double maxPageWidthInPoints,
      OffsetHolder offsets,
      boolean first) {
    double maxHeight = Nvl.nvl(maxPageHeightInPoints, MAX_PAGE_HEIGHT_IN_POINTS);
    double maxWidth = Nvl.nvl(maxPageWidthInPoints, MAX_PAGE_WIDTH_IN_POINTS);
    ColumnSizingData columnSizingData = new ColumnSizingData();
    boolean isFirst = first;
    for (int i = 0; i < sourceWorkbook.getNumberOfSheets(); ++i) {
      Sheet sourceSheet = sourceWorkbook.getSheetAt(i);
      if (sourceSheet.rowIterator().hasNext()) {
        int sourceSheetCellWidth = getMaxColumnNumber(sourceSheet);
        double sourceSheetWidthInPoints =
            getSheetWidthInPoints(sourceSheet, 0, sourceSheetCellWidth);
        offsets.numberInRow++;

        if ((sourceSheetWidthInPoints
                    + getSheetWidthInPoints(targetSheet, null, offsets.columnOffset - 1)
                > maxWidth)
            || (offsets.numberInRow > maxPerRow)
            || forcePageBreaks) {
          // Определение положения новых смещений для текущего элемента
          offsets.rowOffset = getNextUncreatedRowIndex(targetSheet, first);
          offsets.columnOffset = 0;
          offsets.numberInRow = 0;
          offsets.numberInPage++;

          double sourceSheetHeightInPoints = getSheetHeightInPoints(sourceSheet);
          int[] targetRowBreaks = targetSheet.getRowBreaks();
          double targetSheetHeightInPoints =
              getSheetHeightInPoints(
                  targetSheet,
                  targetRowBreaks.length > 0
                      ? (targetRowBreaks[targetRowBreaks.length - 1] + 1)
                      : 0,
                  null);
          if ((offsets.numberInPage == maxPerPage)
              || forcePageBreaks
              || (targetSheetHeightInPoints + sourceSheetHeightInPoints > maxHeight)) {
            offsets.numberInPage = 0;
            targetSheet.setRowBreak(targetSheet.getLastRowNum());
          }
        }

        // Копирование данных, размеров и стилей листа
        for (Row row : sourceSheet) {
          copyRow(
              sourceWorkbook,
              targetSheet,
              targetWorkbook,
              columnSizingData,
              sourceSheet,
              row,
              offsets.rowOffset,
              offsets.columnOffset);
        }
        applyMergedRegions(
            targetSheet,
            shiftRegions(extractRegions(sourceSheet), offsets.rowOffset, offsets.columnOffset));

        offsets.columnOffset += sourceSheetCellWidth;
        isFirst = false;
      }
    }
    return isFirst;
  }

  public static void simpleCollectToSingle(
      Workbook sourceWb, Sheet targetSheet, Workbook targetWb) {
    collectToSingle(
        sourceWb,
        targetSheet,
        targetWb,
        true,
        Integer.MAX_VALUE,
        1,
        null,
        null,
        new OffsetHolder(),
        true);
  }

  /** Получение полного описания объединённых ячеек */
  static List<CellRangeAddress> extractRegions(Sheet sheet) {
    List<CellRangeAddress> result = new ArrayList<>();
    for (int j = 0; j < sheet.getNumMergedRegions(); ++j) {
      result.add(sheet.getMergedRegion(j));
    }
    return result;
  }

  /** Применение описаний объединённых ячеек к листу */
  static void applyMergedRegions(Sheet targetSheet, Collection<CellRangeAddress> regions) {
    for (CellRangeAddress region : regions) {
      targetSheet.addMergedRegion(region);
    }
  }

  public static List<CellRangeAddress> shiftRegions(
      List<CellRangeAddress> regions, Integer rowOffset, Integer columnOffset) {
    int rowOffsetFinal = Nvl.nvl(rowOffset, 0);
    int columnOffsetFinal = Nvl.nvl(columnOffset, 0);
    List<CellRangeAddress> result = new ArrayList<>(regions.size());
    for (CellRangeAddress region : regions) {
      int shiftedFirstRow = Nvl.max(region.getFirstRow() + rowOffsetFinal, 0);
      int shiftedFirstColumn = Nvl.max(region.getFirstColumn() + columnOffsetFinal, 0);
      result.add(
          new CellRangeAddress(
              shiftedFirstRow,
              Nvl.max(shiftedFirstRow, (region.getLastRow() + rowOffsetFinal)),
              shiftedFirstColumn,
              Nvl.max(shiftedFirstColumn, (region.getLastColumn() + columnOffsetFinal))));
    }
    return result;
  }

  /**
   * Получение "исправленного" списка объединённых ячеек при частичном копировании
   *
   * @param regions Список объединённых ячеек
   * @param firstRow Начальная строка
   * @param lastRow Конечная строка
   * @param firstColumn Начальный столбец
   * @param lastColumn Последний столбец
   */
  public static List<CellRangeAddress> filterRegions(
      List<CellRangeAddress> regions,
      Integer firstRow,
      Integer lastRow,
      Integer firstColumn,
      Integer lastColumn) {
    List<CellRangeAddress> result = new ArrayList<>();
    for (CellRangeAddress region : regions) {
      int croppedFirstRow = Nvl.max(firstRow, region.getFirstRow());
      int croppedLastRow = Nvl.min(lastRow, region.getLastRow());
      int croppedFirstColumn = Nvl.max(firstColumn, region.getFirstColumn());
      int croppedLastColumn = Nvl.min(lastColumn, region.getLastColumn());
      if (croppedFirstRow <= croppedLastRow && croppedFirstColumn <= croppedLastColumn) {
        result.add(
            new CellRangeAddress(
                croppedFirstRow, croppedLastRow, croppedFirstColumn, croppedLastColumn));
      }
    }
    return result;
  }

  /**
   * Копирование строки из одного документа в другой, с сохранением значений, стиля, ширины и высоты
   *
   * @param sourceWorkbook Книга источник
   * @param targetSheet Целевой лист
   * @param targetWorkbook Целевая книга
   * @param columnSizingData Структура с информацией о назначении размеров столбцов целевого листа
   * @param sourceSheet Исходный лист
   * @param row Строка исходных данных, копирование которой необходимо осуществить
   */
  static void copyRow(
      Workbook sourceWorkbook,
      Sheet targetSheet,
      Workbook targetWorkbook,
      ColumnSizingData columnSizingData,
      Sheet sourceSheet,
      Row row,
      Integer rowOffset,
      Integer columnOffset) {
    if (row != null) {
      Row targetRow = getOrCreateRow(targetSheet, Nvl.nvl(rowOffset, 0) + row.getRowNum());
      targetRow.setHeightInPoints(row.getHeightInPoints());
      for (Cell cell : row) {
        copyCell(
            sourceWorkbook,
            targetSheet,
            targetWorkbook,
            columnSizingData,
            sourceSheet,
            targetRow,
            cell,
            columnOffset);
      }
    }
  }

  /**
   * Копирование ячейки из одного документа в другой, с сохранением значения, стиля, ширины и высоты
   *
   * @param sourceWorkbook Книга источник
   * @param targetSheet Целевой лист
   * @param targetWorkbook Целевая книга
   * @param columnSizingData Структура с информацией о назначении размеров столбцов целевого листа
   * @param sourceSheet Исходный лист
   * @param targetRow Целевая строка
   * @param cell Ячейка исходных данных, которую необходимо скопировать
   */
  static void copyCell(
      Workbook sourceWorkbook,
      Sheet targetSheet,
      Workbook targetWorkbook,
      ColumnSizingData columnSizingData,
      Sheet sourceSheet,
      Row targetRow,
      Cell cell,
      Integer columnOffset) {
    if (cell != null) {
      int cellColumnIndex = cell.getColumnIndex();
      int targetCellColumnIndex = Nvl.nvl(columnOffset, 0) + cellColumnIndex;
      if (!columnSizingData.columns.contains(targetCellColumnIndex)) {
        columnSizingData.columns.add(targetCellColumnIndex);
        targetSheet.setColumnWidth(
            targetCellColumnIndex, sourceSheet.getColumnWidth(cellColumnIndex));
      }
      if (targetCellColumnIndex > columnSizingData.maxColumnIndex) {
        columnSizingData.maxColumnIndex = targetCellColumnIndex;
      }
      Cell targetCell = getOrCreateCell(targetRow, targetCellColumnIndex);
      if (cell.getCellComment() != null) {
        targetCell.setCellComment(cell.getCellComment());
      }
      CellStyle targetCellStyle = copyStyleTo(sourceWorkbook, cell.getCellStyle(), targetWorkbook);
      if (targetCellStyle != null) {
        targetCell.setCellStyle(targetCellStyle);
      }
      CellType type = cell.getCellType();
      switch (type) {
        case BOOLEAN:
          targetCell.setCellValue(cell.getBooleanCellValue());
          break;
        case NUMERIC:
          targetCell.setCellValue(cell.getNumericCellValue());
          break;
        case STRING:
        case FORMULA:
          targetCell.setCellValue(cell.getNumericCellValue());
          break;
      }
    }
  }

  /**
   * Добавление стилей из одной книги в другую. <br>
   * Функция нужна для обхода бага Excel: <br>
   * при добавлении данных и стилей вперемежку <br>
   * и последующих открытия и сохранения файла в MS Excel в файле слетает таблица стилей
   */
  public static void collectStylesToSingle(Workbook sourceWorkbook, Workbook targetWorkbook) {
    // Шрифты
    for (short i = 0; i < sourceWorkbook.getNumberOfFonts(); ++i) {
      copyFontTo(sourceWorkbook.getFontAt(i), targetWorkbook);
    }
    // Стили
    for (short i = 0; i < sourceWorkbook.getNumCellStyles(); ++i) {
      copyStyleTo(sourceWorkbook, sourceWorkbook.getCellStyleAt(i), targetWorkbook);
    }
  }

  public static byte[] collectToSingle(
      List<byte[]> workbooks,
      Integer maxPerPage,
      Integer maxPerRow,
      boolean forcePageBreak,
      Double maxPageHeightInPoints,
      Double maxPageWidthInPoints)
      throws IOException {
    return collectToSingle(
        workbooks.iterator(),
        maxPerPage,
        maxPerRow,
        forcePageBreak,
        maxPageHeightInPoints,
        maxPageWidthInPoints);
  }

  public static Workbook collectToSingleWorkBook(
      List<byte[]> workbooks,
      Integer maxPerPage,
      Integer maxPerRow,
      boolean forcePageBreak,
      Double maxPageHeightInPoints,
      Double maxPageWidthInPoints)
      throws IOException {
    return collectToSingleWorkBook(
        workbooks.iterator(),
        maxPerPage,
        maxPerRow,
        forcePageBreak,
        maxPageHeightInPoints,
        maxPageWidthInPoints);
  }

  public static Workbook collectToSingleWorkBook(
      byte[][] workbooks,
      Integer maxPerPage,
      Integer maxPerRow,
      boolean forcePageBreak,
      Double maxPageHeightInPoints,
      Double maxPageWidthInPoints)
      throws IOException {
    return collectToSingleWorkBook(
        Arrays.asList(workbooks),
        maxPerPage,
        maxPerRow,
        forcePageBreak,
        maxPageHeightInPoints,
        maxPageWidthInPoints);
  }

  public static byte[] collectToSingle(
      byte[][] workbooks,
      Integer maxPerPage,
      Integer maxPerRow,
      boolean forcePageBreak,
      Double maxPageHeightInPoints,
      Double maxPageWidthInPoints)
      throws IOException {
    return collectToSingle(
        Arrays.asList(workbooks),
        maxPerPage,
        maxPerRow,
        forcePageBreak,
        maxPageHeightInPoints,
        maxPageWidthInPoints);
  }

  public static Workbook collectStylesOnly(Iterator<byte[]> workbooksData) throws IOException {
    return collectStylesOnly(workbooksData, null);
  }

  /**
   * Функция сбора стилей ячеек из набора книг в одну
   *
   * @param workbooksData Список книг из которых нужно получить стили
   * @param workbook Книга в которую будут собраны стили, если книга не передана, то будет создана
   *     новая пустая
   * @return Возвращается модифицированная (или созданная книга) с собранными стилями
   * @throws IOException
   */
  public static Workbook collectStylesOnly(Iterator<byte[]> workbooksData, Workbook workbook)
      throws IOException {
    return collectStylesOnlyInner(workbooksData, workbook == null ? new HSSFWorkbook() : workbook);
  }

  private static Workbook collectStylesOnlyInner(Iterator<byte[]> workbooksData, Workbook workbook)
      throws IOException {
    while (workbooksData.hasNext()) {
      collectStylesToSingle(
          new HSSFWorkbook(new ByteArrayInputStream(workbooksData.next())), workbook);
    }
    return workbook;
  }

  /**
   * Функция осуществляющая соединение содержания нескольких книг в одну
   *
   * @param workbooksData Список книг, содержимое которых необходимо соединить
   * @param maxPerPage Максимальное количество соединяемых элементов на одну страницу итогового
   *     документа
   * @param forcePageBreak Принудительно вставлять перенос страницы в результирующий документ после
   *     каждого добавленного элемента
   * @param maxPageHeightInPoints Максимальная высота страницы результирующего документа, во
   *     избежание превышения которого автоматически вставляется перенос страницы
   * @throws IOException
   */
  public static Workbook collectToSingleWorkBook(
      Iterator<byte[]> workbooksData,
      Integer maxPerPage,
      Integer maxPerRow,
      boolean forcePageBreak,
      Double maxPageHeightInPoints,
      Double maxPageWidthInPoints)
      throws IOException {
    Workbook targetWorkbook = new HSSFWorkbook();
    Sheet targetSheet = targetWorkbook.createSheet("Лист");
    List<Workbook> workbooks = new ArrayList<>();
    while (workbooksData.hasNext()) {
      Workbook sourceWorkbook = new HSSFWorkbook(new ByteArrayInputStream(workbooksData.next()));
      workbooks.add(sourceWorkbook);
      collectStylesToSingle(sourceWorkbook, targetWorkbook);
    }
    OffsetHolder offsets = new OffsetHolder();
    boolean first = true;
    for (Workbook sourceWorkbook : workbooks) {
      first =
          collectToSingle(
              sourceWorkbook,
              targetSheet,
              targetWorkbook,
              forcePageBreak,
              maxPerPage,
              maxPerRow,
              maxPageHeightInPoints,
              maxPageWidthInPoints,
              offsets,
              first);
    }
    return targetWorkbook;
  }

  public static byte[] collectToSingle(
      Iterator<byte[]> workbooks,
      Integer maxPerPage,
      Integer maxPerRow,
      boolean forcePageBreak,
      Double maxPageHeightInPoints,
      Double maxPageWidthInPoints)
      throws IOException {
    return PoiUtils.flushWorkBook(
        collectToSingleWorkBook(
            workbooks,
            maxPerPage,
            maxPerRow,
            forcePageBreak,
            maxPageHeightInPoints,
            maxPageWidthInPoints));
  }

  public static double getSheetHeightInPoints(Sheet sheet) {
    return getSheetHeightInPoints(sheet, null, null);
  }

  /**
   * Определение высоты листа в точках
   *
   * @param sheet Лист, высоту которого нужно измерить
   * @param firstRowIndex Индекс строки, начиная с которой нужно проводить измерение, если передано
   *     null, то нижняя граница номеров строк не учитывается
   * @param lastRowIndex Индекс строки, заканчивая которой нужно проводить измерения, если передано
   *     null, то верхняя граница номеров строк не учитывается
   */
  public static double getSheetHeightInPoints(
      Sheet sheet, Integer firstRowIndex, Integer lastRowIndex) {
    double sum = 0;
    int firstRow = firstRowIndex == null ? 0 : firstRowIndex;
    int lastRow = lastRowIndex == null ? sheet.getLastRowNum() : lastRowIndex;
    for (Row row : sheet) {
      int rowNum = row.getRowNum();
      if (rowNum >= firstRow && rowNum <= lastRow) {
        sum += row.getHeightInPoints();
      }
    }
    return sum;
  }

  public static double getSheetWidthInPoints(Sheet sheet) {
    return getSheetWidthInPoints(sheet, null, null);
  }

  /**
   * Определение ширины листа в точках
   *
   * @param sheet Лист, ширину которого нужно измерить
   * @param firstColumnIndex Индекс столбца, начиная с которого нужно проводить измерение, если
   *     передано null, то нижняя граница номеров столбцов не учитывается
   * @param lastColumnIndex Индекс столбца, заканчивая которым нужно проводить измерения, если
   *     передано null, то верхняя граница номеров столбцов не учитывается
   */
  public static double getSheetWidthInPoints(
      Sheet sheet, Integer firstColumnIndex, Integer lastColumnIndex) {
    double sum = 0;
    int firstColumn = firstColumnIndex == null ? 0 : firstColumnIndex;
    int lastColumn = lastColumnIndex == null ? getMaxColumnNumber(sheet) : lastColumnIndex;
    for (int i = firstColumn; i < lastColumn; ++i) {
      sum += Nvl.nvl(sheet.getColumnWidth(i), 0);
    }
    return sum;
  }

  /**
   * Получение высоты листа
   *
   * @param sheet Лист, высоту которого нужно измерить
   * @return
   */
  public static int getSheetHeight(Sheet sheet) {
    return stream(sheet.spliterator(), false).mapToInt(Row::getHeight).sum();
  }

  /** Автоматическое изменение высоты всех строк так, чтобы текст был полностью виден */
  public static void autoSizeRows(Sheet sheet, Workbook workbook) {
    for (Row row : sheet) {
      SSFExtraTools.autoSizeRow(workbook, sheet, row, true);
    }
  }

  /** Автоматическое изменение ширины интервала столбцов так, чтобы текст был полностью виден */
  public static void autoSizeColumns(
      Sheet sheet, Workbook workbook, int baseColumn, int horizontalLength) {
    for (int i = baseColumn; i < baseColumn + horizontalLength; ++i) {
      SSFExtraTools.autoSizeColumn(workbook, sheet, i, true);
      int width = (sheet.getColumnWidth(i) * 115) / 100;
      sheet.setColumnWidth(i, max(width, MIN_WIDTH));
    }
  }

  /** Получение самого большого индекса столбца на листе */
  public static int getLastColumnNum(Sheet sheet) {
    int result = -1;
    for (Row row : sheet) {
      result = Math.max(row.getLastCellNum(), result);
    }
    return result;
  }

  /** Осуществление получения части листа, ограниченного заданными интервалом строк */
  private static Workbook extractRows(
      Workbook workbook, Sheet sheet, Integer firstRowIndex, Integer lastRowIndex) {
    Workbook targetWorkbook = new HSSFWorkbook();
    collectStylesToSingle(workbook, targetWorkbook);
    Sheet targetSheet = targetWorkbook.createSheet();
    ColumnSizingData columnSizingData = new ColumnSizingData();
    int firstRow = firstRowIndex == null ? 0 : firstRowIndex;
    int lastRow = lastRowIndex == null ? 0 : sheet.getLastRowNum();
    for (int rowIndex = firstRow; rowIndex <= lastRow; ++rowIndex) {
      copyRow(
          workbook,
          targetSheet,
          targetWorkbook,
          columnSizingData,
          sheet,
          sheet.getRow(rowIndex),
          -firstRow,
          0);
    }
    applyMergedRegions(
        targetSheet,
        shiftRegions(
            filterRegions(extractRegions(sheet), firstRow, lastRow, null, null), -firstRow, 0));
    return targetWorkbook;
  }

  private static Workbook checkSplittedItem(Workbook workbook) {
    if (workbook != null) {
      for (int i = 0; i < workbook.getNumberOfSheets(); ++i) {
        for (Row row : workbook.getSheetAt(i)) {
          for (Cell cell : row) {
            if (cell != null) {
              return workbook;
            }
          }
        }
      }
    }
    return null;
  }

  /** Осуществление постраничного разбиения книги на набор отдельных книг */
  public static List<Workbook> splitByPages(Workbook workbook) {
    List<Workbook> result = new ArrayList<>();
    for (int i = 0; i < workbook.getNumberOfSheets(); ++i) {
      Sheet sheet = workbook.getSheetAt(i);
      int currentRowIndex = 0;
      for (int breakRowIndex : sheet.getRowBreaks()) {
        result.add(checkSplittedItem(extractRows(workbook, sheet, currentRowIndex, breakRowIndex)));
        currentRowIndex = breakRowIndex + 1;
      }
      result.add(checkSplittedItem(extractRows(workbook, sheet, currentRowIndex, null)));
    }
    return result.stream().filter(Objects::nonNull).collect(toList());
  }

  /** Получение высоты строки, в которой находится ячейка, если она существует, иначе null */
  public static Float getRowHeightByNamedCell(Workbook workbook, Sheet sheet, String cellName) {
    Cell cell = PoiUtils.getNamedCell(workbook, cellName);
    return cell == null ? null : sheet.getRow(cell.getRowIndex()).getHeightInPoints();
  }

  public static Row getRowByNamedCell(Workbook workbook, Sheet sheet, String cellName) {
    Cell cell = PoiUtils.getNamedCell(workbook, cellName);
    return cell == null ? null : sheet.getRow(cell.getRowIndex());
  }

  /** Remove a row by its index */
  public static void removeRow(Sheet sheet, int rowIndex) {
    int lastRowNum = sheet.getLastRowNum();
    if (rowIndex >= 0 && rowIndex < lastRowNum) {
      sheet.shiftRows(rowIndex + 1, lastRowNum, -1);
    }
    if (rowIndex == lastRowNum) {
      Row removingRow = sheet.getRow(rowIndex);
      if (removingRow != null) {
        sheet.removeRow(removingRow);
      }
    }
  }

  public static void applyMargins(Sheet target, Sheet source) {
    target.getPrintSetup().setPaperSize(source.getPrintSetup().getPaperSize());
    target.getPrintSetup().setLandscape(source.getPrintSetup().getLandscape());
    Stream.of(
            PageMargin.LEFT,
            PageMargin.RIGHT,
            PageMargin.TOP,
            PageMargin.BOTTOM,
            PageMargin.HEADER,
            PageMargin.FOOTER)
        .forEach(margin -> target.setMargin(margin, source.getMargin(margin)));
  }

  public static void cleanRowBreaks(Workbook workbook) {
    workbook.forEach(
        sheet -> {
          int[] breaks = sheet.getRowBreaks();
          if (ArrayUtils.isNotEmpty(breaks) && breaks[0] == 0) {
            sheet.removeRowBreak(0);
          }
        });
  }
}
