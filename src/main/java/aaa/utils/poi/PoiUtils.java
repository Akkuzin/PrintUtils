package aaa.utils.poi;

import static aaa.nvl.Nvl.nvl;
import static com.google.common.base.Strings.nullToEmpty;
import static java.lang.Math.max;
import static java.util.Arrays.asList;
import static java.util.stream.StreamSupport.stream;
import static org.apache.commons.lang3.StringUtils.isNotBlank;

import aaa.nvl.Nvl;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Date;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Set;
import java.util.stream.Stream;
import lombok.AccessLevel;
import lombok.NonNull;
import lombok.SneakyThrows;
import lombok.experimental.FieldDefaults;
import lombok.experimental.UtilityClass;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFFooter;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.SSFExtraTools;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/** Apache POI utils */
@UtilityClass
public class PoiUtils {

  private static final int PERCENT_SAFE_MARGIN = 115;

  private static final int PERCENT_100 = 100;

  private static final int MIN_WIDTH = 256 * 4;

  private static final double MAX_PAGE_HEIGHT_IN_POINTS = 842;
  private static final double MAX_PAGE_WIDTH_IN_POINTS = 27500;
  public static final int ROTATE_RIGHT_XSSF = 180;
  public static final int ROTATE_RIGHT_HSSF = -90;

  /**
   * Get row with certain index, create if non
   *
   * @param sheet Sheet in which row will be
   * @param rowIndex Row index in sheet
   * @return
   */
  public static Row getOrCreateRow(Sheet sheet, int rowIndex) {
    Row row = sheet.getRow(rowIndex);
    if (row == null) {
      row = sheet.createRow(rowIndex);
    }
    return row;
  }

  /**
   * Get cell with certain index, create if non
   *
   * @param row Row in which cell will be
   * @param cellIndex Cell index in row
   * @return
   */
  public static Cell getOrCreateCell(Row row, int cellIndex) {
    Cell cell = row.getCell(cellIndex);
    if (cell == null) {
      cell = row.createCell(cellIndex);
    }
    return cell;
  }

  /**
   * Get cell with certain address, create if non
   *
   * @param sheet Sheet to be updated
   * @param rowIndex Row in which cell will be
   * @param cellIndex Cell index in row
   * @return
   */
  public static Cell getOrCreateCell(Sheet sheet, int rowIndex, int cellIndex) {
    return getOrCreateCell(getOrCreateRow(sheet, rowIndex), cellIndex);
  }

  /** createCustomCell functions family Creates cell at certain coordinates with style and value */
  public static Cell createCustomCell(Row row, long column, CellStyle style, RichTextString value) {
    Cell cell = getOrCreateCell(row, (int) column);
    if (value != null) {
      cell.setCellValue(value);
    }
    if (style != null) {
      cell.setCellStyle(style);
    }
    return cell;
  }

  public static Cell createCustomCell(Row row, long column, CellStyle style, String value) {
    return createCustomCell(
        row,
        column,
        style,
        row.getSheet().getWorkbook().getCreationHelper().createRichTextString(value));
  }

  public static Cell createCustomCell(Row row, long column, CellStyle style, Long value) {
    Cell cell = getOrCreateCell(row, (int) column);
    if (value != null) {
      cell.setCellValue(value);
    }
    if (style != null) {
      cell.setCellStyle(style);
    }
    return cell;
  }

  public static Cell createCustomCell(Row row, long column, CellStyle style, Double value) {
    Cell cell = getOrCreateCell(row, (int) column);
    if (value != null) {
      cell.setCellValue(value);
    }
    if (style != null) {
      cell.setCellStyle(style);
    }
    return cell;
  }

  public static Cell createCustomCell(
      Sheet sheet, long row, long column, CellStyle style, HSSFRichTextString value) {
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
    getOrCreateCell(sheet, (int) row, (int) column).setCellValue((HSSFRichTextString) null);
  }

  /**
   * Flushes workbook to byte array
   *
   * @param workBook Workbook to be flushed
   * @return
   */
  @SneakyThrows
  public static byte[] flushWorkBook(Workbook workBook) {
    ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
    workBook.write(outputStream);
    outputStream.close();
    return outputStream.toByteArray();
  }

  /**
   * Add pages counter to single sheet of workbook
   *
   * @param sheet Sheet to which numbering will be added
   * @param pageWord Word "page" (for internationalization purposes)
   * @param totalPagesDelimiter Delimiter between current page number and total pages count number
   */
  public static void addPagesNumbers(Sheet sheet, String pageWord, String totalPagesDelimiter) {
    sheet
        .getFooter()
        .setRight(
            (isNotBlank(pageWord) ? pageWord + " " : "")
                + HSSFFooter.page()
                + nullToEmpty(totalPagesDelimiter)
                + HSSFFooter.numPages());
  }

  /**
   * Makes named cell in excel workbook
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
   * Makes named cell in excel workbook
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
   * Makes named cell in excel workbook
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
    CellReference cellReference =
        new CellReference(sheetName, (int) rowIndex, (int) columnIndex, false, false);
    namedCell.setNameName(name);
    namedCell.setRefersToFormula(cellReference.formatAsString());
  }

  /**
   * Get reference to named cell
   *
   * @param workbook Workbook object
   * @param cellName Name of named cell
   * @return
   */
  public static CellReference getNamedCellReference(Workbook workbook, String cellName) {
    Name namedCell = workbook.getName(cellName);
    return namedCell == null ? null : new CellReference(namedCell.getRefersToFormula());
  }

  /**
   * Get named cell value
   *
   * @param workbook Workbook object
   * @param cellName Name of named cell
   * @return
   */
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
    setNamedCellValue(
        workbook, cellName, workbook.getCreationHelper().createRichTextString(value), style);
  }

  public static void setNamedCellValue(Workbook workbook, String cellName, String value) {
    setNamedCellValue(workbook, cellName, workbook.getCreationHelper().createRichTextString(value));
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
    int row = -1;
    Name namedCell = workbook.getName(cellName);
    if (namedCell != null) {
      row =
          new AreaReference(namedCell.getRefersToFormula(), workbook.getSpreadsheetVersion())
              .getFirstCell()
              .getRow();
    }
    return row;
  }

  public static int getNamedCellColumnIndex(Workbook workbook, String cellName) {
    int col = -1;
    Name namedCell = workbook.getName(cellName);
    if (namedCell != null) {
      col =
          new AreaReference(namedCell.getRefersToFormula(), workbook.getSpreadsheetVersion())
              .getFirstCell()
              .getCol();
    }
    return col;
  }

  /**
   * Получение крайнего (максимального) номера столбца на листе
   *
   * @param sheet Лист, горизонтальную размерность которого нужно измерить
   * @return
   */
  public static int getMaxColumnNumber(Sheet sheet) {
    return stream(sheet.spliterator(), false).mapToInt(Row::getLastCellNum).max().orElse(0);
  }

  /**
   * Проверка равенства 2х описаний шрифтов
   *
   * @param font1 Первое описание шрифта
   * @param font2 Второе описание шрифта
   * @return
   */
  public static boolean fontsEquals(Font font1, Font font2) {
    return (font1.getItalic() == font2.getItalic())
        && (font1.getStrikeout() == font2.getStrikeout())
        && (font1.getBold() == font2.getBold())
        && (font1.getFontHeightInPoints() == font2.getFontHeightInPoints())
        && (font1.getFontHeight() == font2.getFontHeight())
        && (font1.getColor() == font2.getColor())
        && (StringUtils.equals(font1.getFontName(), font2.getFontName()))
        && (font1.getTypeOffset() == font2.getTypeOffset())
        && (font1.getUnderline() == font2.getUnderline());
  }

  /**
   * Создание копии описания шрифта в другой рабочей книге
   *
   * @param sourceFont Описание шрифта, который нужно скопировать
   * @param targetWorkbook Целевая рабочая книга
   * @return
   */
  public static Font copyFontTo(Font sourceFont, Workbook targetWorkbook, boolean force) {
    if (sourceFont == null) {
      return null;
    }
    if (!force) {
      for (int i = 0; i < targetWorkbook.getNumberOfFontsAsInt(); ++i) {
        Font font = targetWorkbook.getFontAt(i);
        if (fontsEquals(font, sourceFont)) {
          return font;
        }
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

  /**
   * Проверка равенства 2х стилей
   *
   * @param workbook1 Первая книга
   * @param style1 Стиль из первой книги
   * @param workbook2 Вторая книга
   * @param style2 Стиль из второй книги
   * @return
   */
  public static boolean styleEquals(
      Workbook workbook1, CellStyle style1, Workbook workbook2, CellStyle style2) {
    return (style1.getAlignment() == style2.getAlignment())
        && (style1.getVerticalAlignment() == style2.getVerticalAlignment())
        && (style1.getBorderBottom() == style2.getBorderBottom())
        && (style1.getBorderTop() == style2.getBorderTop())
        && (style1.getBorderLeft() == style2.getBorderLeft())
        && (style1.getBorderRight() == style2.getBorderRight())
        && fontsEquals(
            workbook1.getFontAt(style1.getFontIndexAsInt()),
            workbook2.getFontAt(style2.getFontIndexAsInt()))
        // && style1.getFontIndex() == style2.getFontIndex()
        && (style1.getFillBackgroundColor() == style2.getFillBackgroundColor())
        && (style1.getFillForegroundColor() == style2.getFillForegroundColor())
        && (style1.getFillPattern() == style2.getFillPattern())
        && (style1.getIndention() == style2.getIndention())
        && (style1.getRotation() == style2.getRotation())
        && (style1.getWrapText() == style2.getWrapText())
        && StringUtils.equals(
            nullToEmpty(workbook1.createDataFormat().getFormat(style1.getDataFormat())),
            nullToEmpty(workbook2.createDataFormat().getFormat(style2.getDataFormat())));
  }

  public static Workbook workbookFromBytes(@NonNull byte[] data) {
    return workbookFromBytes(new ByteArrayInputStream(data));
  }

  @SneakyThrows
  public static Workbook workbookFromBytes(InputStream data) {
    return WorkbookFactory.create(data);
  }

  @FieldDefaults(level = AccessLevel.PUBLIC)
  public static class OffsetHolder {
    int rowOffset;
    int columnOffset;
    int numberInPage;
    int numberInRow;
  }

  /**
   * Создание копии стиля оформления в другой рабочей книге
   *
   * @param sourceWorkbook Исходная рабочая книга
   * @param sourceStyle Стиль, который нужно скопировать
   * @param targetWorkbook Целевая рабочая книга
   * @return
   */
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
    targetStyle.setVerticalAlignment(sourceStyle.getVerticalAlignment());
    targetStyle.setBorderBottom(sourceStyle.getBorderBottom());
    targetStyle.setBorderTop(sourceStyle.getBorderTop());
    targetStyle.setBorderLeft(sourceStyle.getBorderLeft());
    targetStyle.setBorderRight(sourceStyle.getBorderRight());
    Font targetFont =
        copyFontTo(
            sourceWorkbook.getFontAt(sourceStyle.getFontIndexAsInt()), targetWorkbook, false);
    if (targetFont != null) {
      targetStyle.setFont(targetFont);
    }
    targetStyle.setFillBackgroundColor(sourceStyle.getFillBackgroundColor());
    targetStyle.setFillForegroundColor(sourceStyle.getFillForegroundColor());
    targetStyle.setFillPattern(sourceStyle.getFillPattern());
    targetStyle.setIndention(sourceStyle.getIndention());
    if (targetWorkbook instanceof XSSFWorkbook ^ sourceWorkbook instanceof XSSFWorkbook) {
      if (sourceWorkbook instanceof XSSFWorkbook
          && sourceStyle.getRotation() == ROTATE_RIGHT_XSSF) {
        targetStyle.setRotation((short) ROTATE_RIGHT_HSSF);
      } else if (sourceWorkbook instanceof HSSFWorkbook
          && sourceStyle.getRotation() == ROTATE_RIGHT_HSSF) {
        targetStyle.setRotation((short) ROTATE_RIGHT_XSSF);
      } else {
        targetStyle.setRotation(sourceStyle.getRotation());
      }
    } else {
      targetStyle.setRotation(sourceStyle.getRotation());
    }
    targetStyle.setWrapText(sourceStyle.getWrapText());

    targetStyle.setDataFormat(
        targetWorkbook
            .createDataFormat()
            .getFormat(sourceWorkbook.createDataFormat().getFormat(sourceStyle.getDataFormat())));

    return targetStyle;
  }

  public static int getNextUncreatedRowIndex(Sheet sheet, boolean first) {
    return first ? 0 : sheet.getLastRowNum() + 1;
  }

  @FieldDefaults(level = AccessLevel.PUBLIC)
  private static class ColumnSizingData {
    Set<Integer> columns = new HashSet<>();
    int maxColumnIndex = -1;
  }

  /**
   * Добавление данных из некоторой книги Excel на заданный рабочий лист Excel. Добавление
   * происходит "в столбик". Подразумевается, что данные для в исходной рабочей книге сформированы
   * по одному шаблону
   *
   * @param sourceWorkbook Исходная рабочая книга
   * @param targetSheet Лист на рабочей книге, на котором необходимо свести данные
   * @param targetWorkbook Целевая рабочая книга
   */
  public static boolean collectToSingle(
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
    double maxHeight = nvl(maxPageHeightInPoints, MAX_PAGE_HEIGHT_IN_POINTS);
    double maxWidth = nvl(maxPageWidthInPoints, MAX_PAGE_WIDTH_IN_POINTS);
    ColumnSizingData columnSizingData = new ColumnSizingData();
    for (int i = 0; i < sourceWorkbook.getNumberOfSheets(); ++i) {
      Sheet sourceSheet = sourceWorkbook.getSheetAt(i);
      if (sourceSheet.rowIterator().hasNext()) {

        int sourceSheetCellWidth = getMaxColumnNumber(sourceSheet);
        double sourceSheetWidthInPoints =
            getSheetWidthInPoints(sourceSheet, 0, sourceSheetCellWidth);
        offsets.numberInRow++;

        if (sourceSheetWidthInPoints
                    + getSheetWidthInPoints(targetSheet, null, offsets.columnOffset - 1)
                > maxWidth
            || offsets.numberInRow > maxPerRow
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
          if (offsets.numberInPage == maxPerPage
              || forcePageBreaks
              || targetSheetHeightInPoints + sourceSheetHeightInPoints > maxHeight) {
            offsets.numberInPage = 0;
            if (targetSheet.getLastRowNum() >= 0) {
              targetSheet.setRowBreak(targetSheet.getLastRowNum());
            }
          }
        }

        // Копирование данных, размеров и стилей листа
        for (Row row : sourceSheet) {
          first =
              copyRow(
                  sourceWorkbook,
                  targetSheet,
                  targetWorkbook,
                  first,
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
        first = false;
      }
    }
    //    for (int columnIndex : targetSheet.getColumnBreaks()) {
    //      targetSheet.removeColumnBreak((short) columnIndex);
    //    }
    // targetSheet.setColumnBreak((short) (columnSizingData.maxColumnIndex));
    return first;
  }

  public static void simpleCollectToSingle(
      Workbook sourceWorkbook, Sheet targetSheet, Workbook targetWorkbook) {
    collectToSingle(
        sourceWorkbook,
        targetSheet,
        targetWorkbook,
        true,
        Integer.MAX_VALUE,
        1,
        null,
        null,
        new OffsetHolder(),
        true);
  }

  /**
   * Получение полного описания объединённых ячеек
   *
   * @param sheet Лист, из которого нужно получить описание
   * @return
   */
  protected static List<CellRangeAddress> extractRegions(Sheet sheet) {
    return new ArrayList<>(sheet.getMergedRegions());
  }

  /**
   * Применение описаний объединённых ячеек к листу
   *
   * @param targetSheet Лист, к которому применяется объединение ячеек
   * @param regions Список описаний объединённых ячеек
   */
  protected static void applyMergedRegions(
      Sheet targetSheet, Collection<CellRangeAddress> regions) {
    regions.forEach(targetSheet::addMergedRegionUnsafe);
  }

  public static List<CellRangeAddress> shiftRegions(
      List<CellRangeAddress> regions, Integer rowOffset, Integer columnOffset) {
    int rowOffsetFinal = nvl(rowOffset, 0);
    int columnOffsetFinal = nvl(columnOffset, 0);
    List<CellRangeAddress> result = new ArrayList<>();
    for (CellRangeAddress region : regions) {
      int shiftedFirstRow = max(region.getFirstRow() + rowOffsetFinal, 0);
      int shiftedLastRow = max(shiftedFirstRow, region.getLastRow() + rowOffsetFinal);
      int shiftedFirstColumn = max(region.getFirstColumn() + columnOffsetFinal, 0);
      int shiftedLastColumn = max(shiftedFirstColumn, region.getLastColumn() + columnOffsetFinal);
      result.add(
          new CellRangeAddress(
              shiftedFirstRow, shiftedLastRow, shiftedFirstColumn, shiftedLastColumn));
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
   * @return
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
   * Осуществление копирования строки из одного документа в другой, с сохранением значений, стиля
   * ширины и высоты
   *
   * @param sourceWorkbook Книга источник
   * @param targetSheet Целевой лист
   * @param targetWorkbook Целевая книга
   * @param first Признак того, что копируемая строчка является первой в целевом документе
   * @param columnSizingData Структура содержащая информацию о назначении размеров столбцов целевого
   *     листа
   * @param sourceSheet Исходный лист
   * @param row Строка исходных данных, копирование которой необходимо осуществить
   * @return
   */
  protected static boolean copyRow(
      Workbook sourceWorkbook,
      Sheet targetSheet,
      Workbook targetWorkbook,
      boolean first,
      ColumnSizingData columnSizingData,
      Sheet sourceSheet,
      Row row,
      Integer rowOffset,
      Integer columnOffset) {
    if (row != null) {
      Row targetRow = getOrCreateRow(targetSheet, nvl(rowOffset, 0) + row.getRowNum());
      targetRow.setHeightInPoints(row.getHeightInPoints());
      CellStyle targetRowStyle = copyStyleTo(sourceWorkbook, row.getRowStyle(), targetWorkbook);
      if (targetRowStyle != null) {
        targetRow.setRowStyle(targetRowStyle);
      }
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
    return false;
  }

  /**
   * Осуществление копирования ячейки из одного документа в другой, с сохранением значения, стиля,
   * ширины и высоты
   *
   * @param sourceWorkbook Книга источник
   * @param targetSheet Целевой лист
   * @param targetWorkbook Целевая книга
   * @param columnSizingData Структура содержащая информацию о назначении размеров столбцов целевого
   *     листа
   * @param sourceSheet Исходный лист
   * @param targetRow Целевая строка
   * @param cell Ячейка исходных данных, которую необходимо скопировать
   */
  protected static void copyCell(
      Workbook sourceWorkbook,
      Sheet targetSheet,
      Workbook targetWorkbook,
      ColumnSizingData columnSizingData,
      Sheet sourceSheet,
      Row targetRow,
      Cell cell,
      Integer columnOffset) {
    if (cell != null) {
      int k = cell.getColumnIndex();
      int targetCellColumnIndex = nvl(columnOffset, 0) + k;
      if (!columnSizingData.columns.contains(targetCellColumnIndex)) {
        columnSizingData.columns.add(targetCellColumnIndex);
        targetSheet.setColumnWidth(targetCellColumnIndex, sourceSheet.getColumnWidth(k));
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
      targetCell.setCellType(type);
      if (type == CellType.BOOLEAN) {
        targetCell.setCellValue(cell.getBooleanCellValue());
      } else if (type == CellType.NUMERIC) {
        targetCell.setCellValue(cell.getNumericCellValue());
      } else if (type == CellType.STRING) {
        targetCell.setCellValue(
            targetWorkbook
                .getCreationHelper()
                .createRichTextString(cell.getRichStringCellValue().getString()));
      } else if (type == CellType.FORMULA) {
        targetCell.setCellValue(
            targetWorkbook
                .getCreationHelper()
                .createRichTextString(cell.getRichStringCellValue().getString()));
      }
    }
  }

  /**
   * Добавление стилей из одной книги в другую Функция нужна для обхода бага Excel: при добавлении
   * данных и стилей вперемежку и последующих открытия и сохранения файла в MS Excel в файле слетает
   * таблица стилей
   *
   * @param sourceWorkbook Исходная рабочая книга
   * @param targetWorkbook Целевая рабочая книга
   */
  public static void collectStylesToSingle(Workbook sourceWorkbook, Workbook targetWorkbook) {
    for (int i = 0; i < sourceWorkbook.getNumberOfFontsAsInt(); ++i) {
      copyFontTo(sourceWorkbook.getFontAt(i), targetWorkbook, false);
    }
    for (short i = 0; i < sourceWorkbook.getNumCellStyles(); ++i) {
      copyStyleTo(sourceWorkbook, sourceWorkbook.getCellStyleAt(i), targetWorkbook);
    }
  }

  public static byte[] collectToSingle(
      Collection<byte[]> workbooks,
      Integer maxPerPage,
      Integer maxPerRow,
      boolean forcePageBreak,
      Double maxPageHeightInPoints,
      Double maxPageWidthInPoints) {
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
      Double maxPageWidthInPoints) {
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
      Double maxPageWidthInPoints) {
    return collectToSingleWorkBook(
        asList(workbooks),
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
      Double maxPageWidthInPoints) {
    return collectToSingle(
        asList(workbooks),
        maxPerPage,
        maxPerRow,
        forcePageBreak,
        maxPageHeightInPoints,
        maxPageWidthInPoints);
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
   * @return
   */
  @SneakyThrows
  public static Workbook collectToSingleWorkBook(
      Iterator<byte[]> workbooksData,
      Integer maxPerPage,
      Integer maxPerRow,
      boolean forcePageBreak,
      Double maxPageHeightInPoints,
      Double maxPageWidthInPoints) {

    boolean first = true;
    List<Workbook> workbooks = new ArrayList<>();
    while (workbooksData.hasNext()) {
      workbooks.add(workbookFromBytes(workbooksData.next()));
    }
    Workbook targetWorkbook =
        workbooks.stream()
                .findFirst()
                .map(workbook -> workbook instanceof XSSFWorkbook)
                .orElse(true)
            ? new XSSFWorkbook()
            : new HSSFWorkbook();
    Sheet targetSheet = targetWorkbook.createSheet("Лист");
    workbooks.forEach(sourceWorkbook -> collectStylesToSingle(sourceWorkbook, targetWorkbook));
    OffsetHolder offsets = new OffsetHolder();
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
      Double maxPageWidthInPoints) {
    return flushWorkBook(
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
   * @return
   */
  public static double getSheetHeightInPoints(
      Sheet sheet, Integer firstRowIndex, Integer lastRowIndex) {
    double sum = 0;
    if (firstRowIndex == null) {
      firstRowIndex = 0;
    }
    if (lastRowIndex == null) {
      lastRowIndex = sheet.getLastRowNum();
    }
    for (Row row : sheet) {
      int rowNum = row.getRowNum();
      if (rowNum >= firstRowIndex && rowNum <= lastRowIndex) {
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
   * @return
   */
  public static double getSheetWidthInPoints(
      Sheet sheet, Integer firstColumnIndex, Integer lastColumnIndex) {
    double sum = 0;
    if (firstColumnIndex == null) {
      firstColumnIndex = 0;
    }
    if (lastColumnIndex == null) {
      lastColumnIndex = getMaxColumnNumber(sheet);
    }
    for (int i = firstColumnIndex; i < lastColumnIndex; ++i) {
      sum += nvl(sheet.getColumnWidth(i), 0);
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

  /**
   * Процедура выполняет автоматическое изменение высоты всех строк на листе таким образом, чтобы
   * текст был полностью виден
   *
   * @param sheet Лист. на котором требуется выполнить автоматическое изменение высоты строк
   * @param workbook
   */
  public static void autoSizeRows(Sheet sheet, Workbook workbook) {
    for (Row row : sheet) {
      SSFExtraTools.autoSizeRow(workbook, sheet, row, true);
    }
  }

  /**
   * Процедура выполняет автоматическое изменение ширины интервала столбцов на листе таким образом,
   * чтобы текст был полностью виден
   *
   * @param sheet Лист. на котором требуется выполнить автоматическое изменение высоты строк
   * @param workbook
   */
  public static void autoSizeColumns(
      Sheet sheet, Workbook workbook, int baseColumn, int horizontalLength) {
    for (int i = baseColumn; i < baseColumn + horizontalLength; ++i) {
      SSFExtraTools.autoSizeColumn(workbook, sheet, i, true);
      int width = (sheet.getColumnWidth(i) * PERCENT_SAFE_MARGIN) / PERCENT_100;
      sheet.setColumnWidth(i, max(width, MIN_WIDTH));
    }
  }

  /**
   * Получение самого большого индекса столбца на листе
   *
   * @param sheet Лист, для которого необходимо найти наибольший индекс столбца
   * @return
   */
  public static int getLastColumnNum(Sheet sheet) {
    return stream(sheet.spliterator(), false).mapToInt(Row::getLastCellNum).max().orElse(-1);
  }

  /**
   * Осуществление получения части листа, ограниченного заданными интервалом строк
   *
   * @param workbook Книга к которой принадлежит лист
   * @param sheet Лист, из которого будет скопирована часть строк
   * @param firstRowIndex Индекс первой строки в диапазоне
   * @param lastRowIndex Индекс последней строки в диапазоне
   * @return
   */
  private static Workbook extractRows(
      Workbook workbook, Sheet sheet, Integer firstRowIndex, Integer lastRowIndex) {
    Workbook targetWorkbook =
        workbook instanceof XSSFWorkbook ? new XSSFWorkbook() : new HSSFWorkbook();
    collectStylesToSingle(workbook, targetWorkbook);
    Sheet targetSheet = targetWorkbook.createSheet();
    ColumnSizingData columnSizingData = new ColumnSizingData();
    boolean first = true;
    if (firstRowIndex == null) {
      firstRowIndex = 0;
    }
    if (lastRowIndex == null) {
      lastRowIndex = sheet.getLastRowNum();
    }
    for (int rowIndex = firstRowIndex; rowIndex <= lastRowIndex; ++rowIndex) {
      first =
          copyRow(
              workbook,
              targetSheet,
              targetWorkbook,
              first,
              columnSizingData,
              sheet,
              sheet.getRow(rowIndex),
              -firstRowIndex,
              0);
    }
    applyMergedRegions(
        targetSheet,
        shiftRegions(
            filterRegions(extractRegions(sheet), firstRowIndex, lastRowIndex, null, null),
            -firstRowIndex,
            0));
    return targetWorkbook;
  }

  private static void addSplittedItem(Workbook workbook, List<Workbook> result) {
    if (workbook != null) {
      sheets:
      for (int i = 0; i < workbook.getNumberOfSheets(); ++i) {
        for (Row row : workbook.getSheetAt(i)) {
          for (Cell cell : row) {
            if (cell != null) {
              result.add(workbook);
              break sheets;
            }
          }
        }
      }
    }
  }

  /**
   * Осуществление постраничного разбиения книги на набор отдельных книг
   *
   * @param workbook Рабочая книга, разбиение которой нужно осуществить
   * @return
   */
  public static List<Workbook> splitByPages(Workbook workbook) {
    List<Workbook> result = new ArrayList<>();
    for (int i = 0; i < workbook.getNumberOfSheets(); ++i) {
      Sheet sheet = workbook.getSheetAt(i);
      int currentRowIndex = 0;
      for (int breakRowIndex : sheet.getRowBreaks()) {
        addSplittedItem(extractRows(workbook, sheet, currentRowIndex, breakRowIndex), result);
        currentRowIndex = breakRowIndex + 1;
      }
      addSplittedItem(extractRows(workbook, sheet, currentRowIndex, null), result);
    }
    return result;
  }

  /**
   * @param workbook
   * @param sheet
   * @param cellName
   * @return высоту строки в которой находится ячейка, если она существует или -1
   */
  public static float getRowHeightByNamedCell(Workbook workbook, Sheet sheet, String cellName) {
    Cell cell = getNamedCell(workbook, cellName);
    return cell == null ? -1F : sheet.getRow(cell.getRowIndex()).getHeightInPoints();
  }

  public static Row getRowByNamedCell(Workbook workbook, Sheet sheet, String cellName) {
    Cell cell = getNamedCell(workbook, cellName);
    return cell == null ? null : sheet.getRow(cell.getRowIndex());
  }

  public static Row copyRow(
      Workbook workbook, Sheet worksheet, int sourceRowNum, int destinationRowNum) {
    // Get the source / new row
    Row newRow = worksheet.getRow(destinationRowNum);
    Row sourceRow = worksheet.getRow(sourceRowNum);

    // If the row exist in destination, push down all rows by 1 else create a new row
    if (newRow != null) {
      worksheet.shiftRows(destinationRowNum, worksheet.getLastRowNum(), 1);
    } else {
      newRow = worksheet.createRow(destinationRowNum);
    }

    // Loop through source columns to add to new row
    for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
      // Grab a copy of the old/new cell
      Cell oldCell = sourceRow.getCell(i);
      Cell newCell = newRow.createCell(i);

      // If the old cell is null jump to next cell
      if (oldCell == null) {
        newCell = null;
        continue;
      }

      // Copy style from old cell and apply to new cell
      CellStyle newCellStyle = workbook.createCellStyle();
      newCellStyle.cloneStyleFrom(oldCell.getCellStyle());
      newCell.setCellStyle(newCellStyle);

      // If there is a cell comment, copy
      if (oldCell.getCellComment() != null) {
        newCell.setCellComment(oldCell.getCellComment());
      }

      // If there is a cell hyperlink, copy
      if (oldCell.getHyperlink() != null) {
        newCell.setHyperlink(oldCell.getHyperlink());
      }

      // Set the cell data type
      newCell.setCellType(oldCell.getCellType());

      // Set the cell data value
      switch (oldCell.getCellType()) {
        case BLANK:
          newCell.setBlank();
          break;
        case BOOLEAN:
          newCell.setCellValue(oldCell.getBooleanCellValue());
          break;
        case ERROR:
          newCell.setCellErrorValue(oldCell.getErrorCellValue());
          break;
        case FORMULA:
          newCell.setCellFormula(oldCell.getCellFormula());
          break;
        case NUMERIC:
          newCell.setCellValue(oldCell.getNumericCellValue());
          break;
        case STRING:
          newCell.setCellValue(oldCell.getRichStringCellValue());
          break;
        default:
      }
    }

    // If there are are any merged regions in the source row, copy to new row
    for (CellRangeAddress region : worksheet.getMergedRegions()) {
      if (region.getFirstRow() == sourceRow.getRowNum()) {
        worksheet.addMergedRegion(
            new CellRangeAddress(
                newRow.getRowNum(),
                newRow.getRowNum() + (region.getLastRow() - region.getFirstRow()),
                region.getFirstColumn(),
                region.getLastColumn()));
      }
    }
    return newRow;
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
            Sheet.LeftMargin,
            Sheet.RightMargin,
            Sheet.TopMargin,
            Sheet.BottomMargin,
            Sheet.HeaderMargin,
            Sheet.FooterMargin)
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
