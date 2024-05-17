package aaa.utils.poi;

import static aaa.lang.reflection.IntrospectionUtils.asGetter;
import static aaa.nvl.Nvl.nvl;
import static aaa.nvl.Nvl.nvlToString;
import static aaa.utils.poi.ExcelExport.Alignment.CENTER;
import static aaa.utils.poi.ExcelExport.Alignment.LEFT;
import static aaa.utils.poi.ExcelExport.Alignment.RIGHT;
import static aaa.utils.poi.PoiUtils.autoSizeColumns;
import static aaa.utils.poi.PoiUtils.autoSizeRows;
import static aaa.utils.poi.PoiUtils.flushWorkBook;
import static aaa.utils.poi.PoiUtils.workbookFromBytes;
import static aaa.utils.print.MapResourceBundle.deHyphenate;
import static org.apache.commons.lang3.StringUtils.replaceChars;
import static org.apache.poi.ss.usermodel.PrintSetup.A4_PAPERSIZE;

import jakarta.persistence.Temporal;
import jakarta.persistence.metamodel.SingularAttribute;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Collection;
import java.util.Date;
import java.util.List;
import java.util.Objects;
import java.util.function.Function;
import lombok.AccessLevel;
import lombok.Builder;
import lombok.Builder.Default;
import lombok.Getter;
import lombok.NonNull;
import lombok.Setter;
import lombok.experimental.FieldDefaults;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.SSFExtraTools;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelExport {

  private static final int MAX_PAGE_WIDTH = 31000;
  static final int DEFAULT_FONT_HEIGHT = 10;

  public enum Alignment {
    CENTER,
    LEFT,
    RIGHT
  }

  @FieldDefaults(level = AccessLevel.PUBLIC)
  @Builder
  @Getter
  @Setter
  public static class CellDefinition<T, F> {
    Alignment alignment;
    Integer weight;
    @NonNull String title;
    boolean weightStyle;
    boolean sumStyle;
    boolean dateWithTime;
    @NonNull Function<T, F> accessor;

    public static <T, F> CellDefinitionBuilder<T, F> forField(SingularAttribute<T, F> attribute) {
      return CellDefinition.<T, F>builder()
          .accessor(asGetter(attribute))
          .title(attribute.getName())
          .alignment(
              Temporal.class.isAssignableFrom(attribute.getJavaType())
                  ? CENTER
                  : Number.class.isAssignableFrom(attribute.getJavaType()) ? RIGHT : LEFT);
    }

    public static <T, F> CellDefinitionBuilder<T, F> forGetter(Function<T, F> getter) {
      return CellDefinition.<T, F>builder().accessor(getter);
    }
  }

  @FieldDefaults(level = AccessLevel.PUBLIC)
  @Builder
  @Getter
  @Setter
  public static class PageDefinition {
    @Default short paperSize = PrintSetup.A4_ROTATED_PAPERSIZE;
    @Default String sheetName = "Список";
    @Default boolean landscape = true;
    @Default String fontName = "Arial";
    Integer defaultFontSize;
    @Default boolean prepareTitle = true;
    @Default boolean autoSizeRows = true;
    @Default boolean autoSizeHeader = true;
    Integer maxPageWidth;
  }

  public static byte[] makeReport(Collection<?> data, List<CellDefinition> fields) {
    return makeReport(data, fields, PageDefinition.builder().build());
  }

  public static byte[] makeReport(
      Collection<?> data, List<CellDefinition> fields, PageDefinition pageDefinition) {
    Workbook workbook = new XSSFWorkbook();
    Sheet sheet = workbook.createSheet(pageDefinition.getSheetName());

    CreationHelper createHelper = workbook.getCreationHelper();

    CellStyle titleStyle = workbook.createCellStyle();
    titleStyle.setAlignment(HorizontalAlignment.CENTER);
    titleStyle.setVerticalAlignment(VerticalAlignment.TOP);
    Font fontBold = workbook.createFont();
    fontBold.setFontHeightInPoints(
        nvl(pageDefinition.getDefaultFontSize(), DEFAULT_FONT_HEIGHT).shortValue());
    fontBold.setBold(true);
    fontBold.setFontName(pageDefinition.getFontName());
    titleStyle.setFont(fontBold);
    titleStyle.setWrapText(true);

    Font dataFont = workbook.createFont();
    dataFont.setFontName(pageDefinition.getFontName());
    dataFont.setFontHeightInPoints(
        nvl(pageDefinition.getDefaultFontSize(), DEFAULT_FONT_HEIGHT).shortValue());

    CellStyle leftStyle = workbook.createCellStyle();
    leftStyle.setFont(dataFont);
    leftStyle.setAlignment(HorizontalAlignment.LEFT);
    leftStyle.setWrapText(true);

    CellStyle centeredStyle = workbook.createCellStyle();
    centeredStyle.setFont(dataFont);
    centeredStyle.setAlignment(HorizontalAlignment.CENTER);
    centeredStyle.setWrapText(true);

    CellStyle rightStyle = workbook.createCellStyle();
    rightStyle.setFont(dataFont);
    rightStyle.setAlignment(HorizontalAlignment.RIGHT);
    rightStyle.setWrapText(false);

    CellStyle sumStyle = workbook.createCellStyle();
    sumStyle.setFont(dataFont);
    sumStyle.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.00"));
    sumStyle.setAlignment(HorizontalAlignment.RIGHT);
    sumStyle.setWrapText(false);

    CellStyle weightStyle = workbook.createCellStyle();
    weightStyle.setFont(dataFont);
    weightStyle.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.000"));
    weightStyle.setAlignment(HorizontalAlignment.RIGHT);
    weightStyle.setWrapText(false);

    CellStyle dateStyle = workbook.createCellStyle();
    dateStyle.setFont(dataFont);
    dateStyle.setAlignment(HorizontalAlignment.CENTER);
    dateStyle.setWrapText(false);
    dateStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd.mm.yyyy"));

    CellStyle dateTimeStyle = workbook.createCellStyle();
    dateTimeStyle.setFont(dataFont);
    dateTimeStyle.setAlignment(HorizontalAlignment.CENTER);
    dateTimeStyle.setWrapText(false);
    dateTimeStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd.mm.yyyy hh:mm"));

    int headerRow;

    int rowCounter = 0;
    {
      headerRow = rowCounter;
      Row row = sheet.createRow(rowCounter++);
      int column = 0;
      for (CellDefinition field : fields) {
        Cell cell = row.createCell(column++);
        cell.setCellValue(
            workbook
                .getCreationHelper()
                .createRichTextString(
                    pageDefinition.prepareTitle
                        ? deHyphenate(replaceChars(field.title, "\n\r", "  "))
                        : field.title));
        cell.setCellStyle(titleStyle);
      }
    }

    for (Object bean : data) {
      Row row = sheet.createRow(rowCounter++);
      int column = 0;
      for (CellDefinition field : fields) {
        Cell cell = row.createCell(column++);
        Object value = field.accessor.apply(bean);
        if (value instanceof Number numberValue) {
          cell.setCellValue(numberValue.doubleValue());
          cell.setCellStyle(
              field.sumStyle ? sumStyle : field.weightStyle ? weightStyle : rightStyle);
        } else if (value instanceof Date dateValue) {
          cell.setCellValue(dateValue);
          cell.setCellStyle(field.dateWithTime ? dateTimeStyle : dateStyle);
        } else if (value instanceof LocalDate localDateValue) {
          cell.setCellValue(localDateValue);
          cell.setCellStyle(field.dateWithTime ? dateTimeStyle : dateStyle);
        } else if (value instanceof LocalDateTime localDateTimeValue) {
          cell.setCellValue(localDateTimeValue);
          cell.setCellStyle(field.dateWithTime ? dateTimeStyle : dateStyle);
        } else {
          cell.setCellValue(createHelper.createRichTextString(nvlToString(value)));
          cell.setCellStyle(
              field.alignment == CENTER
                  ? centeredStyle
                  : field.alignment == RIGHT ? rightStyle : leftStyle);
        }
      }
    }

    // Установка размеров столбцов
    {
      double coefficient =
          nvl(pageDefinition.maxPageWidth, MAX_PAGE_WIDTH)
              / fields.stream()
                  .map(CellDefinition::getWeight)
                  .filter(Objects::nonNull)
                  .reduce((i1, i2) -> i1 + i2)
                  .orElse(fields.size());
      sheet.setHorizontallyCenter(true);
      sheet.getPrintSetup().setPaperSize(nvl(pageDefinition.paperSize, A4_PAPERSIZE));
      sheet.getPrintSetup().setLandscape(pageDefinition.landscape);
      int column = 0;
      for (CellDefinition field : fields) {
        sheet.setColumnWidth(column++, (int) Math.ceil(nvl(field.weight, 1) * coefficient));
      }
    }

    if (pageDefinition.autoSizeRows) {
      autoSizeRows(sheet, workbook);
    }
    if (pageDefinition.autoSizeHeader) {
      SSFExtraTools.autoSizeRow(workbook, sheet, sheet.getRow(headerRow), true);
    }

    sheet.setAutoFilter(new CellRangeAddress(0, 0, 0, fields.size() - 1));
    sheet.createFreezePane(0, 1);

    return flushWorkBook(workbook);
  }

  public static byte[] autoExtendColumns(byte[] report, List<ExcelExport.CellDefinition> fields) {
    Workbook workbook = workbookFromBytes(report);
    autoSizeColumns(workbook.iterator().next(), workbook, 0, fields.size());
    autoSizeRows(workbook.iterator().next(), workbook);
    return flushWorkBook(workbook);
  }
}
