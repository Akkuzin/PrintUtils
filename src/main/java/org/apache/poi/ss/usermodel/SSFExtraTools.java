package org.apache.poi.ss.usermodel;

import static org.apache.commons.lang3.StringUtils.isEmpty;
import static org.apache.commons.lang3.StringUtils.split;

import java.awt.font.FontRenderContext;
import java.awt.font.LineBreakMeasurer;
import java.awt.font.TextAttribute;
import java.awt.font.TextLayout;
import java.awt.geom.AffineTransform;
import java.awt.geom.Rectangle2D;
import java.text.AttributedString;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import lombok.AccessLevel;
import lombok.Builder;
import lombok.experimental.FieldDefaults;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.ss.util.CellRangeAddress;

// CHECKSTYLE:OFF

public class SSFExtraTools {

  private static void copyAttributes(Font font, AttributedString str, int startIdx, int endIdx) {
    str.addAttribute(TextAttribute.FAMILY, font.getFontName(), startIdx, endIdx);
    str.addAttribute(TextAttribute.SIZE, Float.valueOf(font.getFontHeightInPoints()));
    if (font.getBold())
      str.addAttribute(TextAttribute.WEIGHT, TextAttribute.WEIGHT_BOLD, startIdx, endIdx);
    if (font.getItalic())
      str.addAttribute(TextAttribute.POSTURE, TextAttribute.POSTURE_OBLIQUE, startIdx, endIdx);
    if (font.getUnderline() == Font.U_SINGLE)
      str.addAttribute(TextAttribute.UNDERLINE, TextAttribute.UNDERLINE_ON, startIdx, endIdx);
  }

  private static boolean containsCell(CellRangeAddress cellRange, int rowIndex, int colIndex) {
    return cellRange.getFirstRow() <= rowIndex
        && cellRange.getLastRow() >= rowIndex
        && cellRange.getFirstColumn() <= colIndex
        && cellRange.getLastColumn() >= colIndex;
  }

  /**
   * Adjusts the row height to fit the contents. This process can be relatively slow on large
   * sheets, so this should normally only be called once per row, at the end of your processing. You
   * can specify whether the content of merged cells should be considered or ignored. Default is to
   * ignore merged cells.
   *
   * @param useMergedCells whether to use the contents of merged cells when calculating the width of
   *     the column
   */
  public static Double calculateRowSize(
      Workbook workbook, Sheet sheet, Row row, boolean useMergedCells) {
    if (workbook == null || sheet == null || row == null) {
      return null;
    }

    double height = -1;
    for (Cell cell : row) {
      if (cell == null) {
        continue;
      }
      int cellIndex = cell.getColumnIndex();

      int widthInChars = sheet.getColumnWidth(cell.getColumnIndex());
      int mergedHeight = 1;
      if (useMergedCells) {
        for (CellRangeAddress region : sheet.getMergedRegions()) {
          if (containsCell(region, row.getRowNum(), cellIndex)) {
            cell = sheet.getRow(region.getFirstRow()).getCell(region.getFirstColumn());
            widthInChars = 0;
            for (int j = region.getFirstColumn(); j <= region.getLastColumn(); ++j) {
              widthInChars += sheet.getColumnWidth(j);
            }
            mergedHeight = region.getLastRow() - region.getFirstRow() + 1;
            break;
          }
        }
      }

      Font font = workbook.getFontAt(cell.getCellStyle().getFontIndex());
      Dimension charDimensions = getCharDimensions(TEST_STRING_H, font);
      // etalon_h[1] = 16;
      charDimensions.height *= 1.7;
      widthInChars /= 200;
      double newHeight = 0;
      if (cell.getCellType() == CellType.STRING) {
        RichTextString rt = cell.getRichStringCellValue();
        String value = rt.getString();
        newHeight =
            getStringLinesCount(
                    split(value, "\n\r"), font, (float) (widthInChars * charDimensions.width))
                * charDimensions.height;
      }
      height = Math.max(height, newHeight / mergedHeight);
    }
    if (height >= 0) {
      // height *= 256;
      // width can be bigger that Short.MAX_VALUE!
      return Math.min(height, Short.MAX_VALUE);
    }
    return null;
  }

  public static Double calculateRowSize(
      Workbook workbook, Sheet sheet, int rowIndex, boolean useMergedCells) {
    return calculateRowSize(workbook, sheet, sheet.getRow(rowIndex), useMergedCells);
  }

  public static void autoSizeRow(Workbook workbook, Sheet sheet, Row row, boolean useMergedCells) {
    Double height = calculateRowSize(workbook, sheet, row, useMergedCells);
    if (height != null) {
      row.setHeightInPoints(height.floatValue());
    }
  }

  public static void autoSizeRow(
      Workbook workbook, Sheet sheet, int rowIndex, boolean useMergedCells) {
    autoSizeRow(workbook, sheet, sheet.getRow(rowIndex), useMergedCells);
  }

  private static final String TEST_STRING_H =
      "asdaklsjdklsajd|wwwwwwwwwwwwwwwwwwwww!@#$%^&*()lkjsalkdjlsakjdlsakqiowuiobv[]z1234567890_+-=";

  private static final String TEST_STRING_V =
      "LineBreakMeasurer is constructed with an iterator over styled text. "
          + "The iterator's range should be a single paragraph in the text.";

  @Builder
  @FieldDefaults(level = AccessLevel.PUBLIC)
  public static class Dimension {
    double width;
    double height;
  }

  private static Dimension getCharDimensions(String string, Font font) {
    Dimension stringSizes = getStringSizes(string, font);
    return Dimension.builder()
        .width(stringSizes.width / string.length())
        .height(stringSizes.height)
        .build();
  }

  private static Dimension getStringSizes(String string, Font font) {
    AttributedString str = new AttributedString(string);
    FontRenderContext frc = new FontRenderContext(null, true, true);
    copyAttributes(font, str, 0, string.length() - 1);
    TextLayout layout = new TextLayout(str.getIterator(), frc);
    Rectangle2D bounds = layout.getBounds();
    return Dimension.builder().width(bounds.getWidth()).height(bounds.getHeight()).build();
  }

  private static int getStringLinesCount(String[] strings, Font font, float maxWidth) {
    int result = 0;
    for (String string : strings) {
      result += getStringLinesCount(string, font, maxWidth);
    }
    return result;
  }

  private static int getStringLinesCount(String string, Font font, float maxWidth) {
    if (isEmpty(string) || string.length() <= 1) {
      return 1;
    }
    AttributedString str = new AttributedString(string);
    FontRenderContext frc = new FontRenderContext(null, true, true);
    copyAttributes(font, str, 0, string.length() - 1);
    LineBreakMeasurer measurer = new LineBreakMeasurer(str.getIterator(), frc);
    int result = 0;
    while (measurer.nextLayout(maxWidth) != null) {
      result++;
    }
    return result;
  }

  /**
   * Adjusts the column width to fit the contents. This process can be relatively slow on large
   * sheets, so this should normally only be called once per column, at the end of your processing.
   * You can specify whether the content of merged cells should be considered or ignored. Default is
   * to ignore merged cells.
   *
   * @param column the column index
   * @param useMergedCells whether to use the contents of merged cells when calculating the width of
   *     the column
   */
  public static void autoSizeColumn(
      Workbook book, Sheet sheet, int column, boolean useMergedCells) {
    AttributedString str;
    TextLayout layout;
    /**
     * Excel measures columns in units of 1/256th of a character width but the docs say nothing
     * about what particular character is used. '0' looks to be a good choice.
     */
    char defaultChar = '0';

    /**
     * This is the multiple that the font height is scaled by when determining the boundary of
     * rotated text.
     */
    double fontHeightMultiple = 2.0;

    FontRenderContext frc = new FontRenderContext(null, true, true);

    // TODO
    Font defaultFont = book.getFontAt(0);

    str = new AttributedString("" + defaultChar);
    copyAttributes(defaultFont, str, 0, 1);
    layout = new TextLayout(str.getIterator(), frc);
    int defaultCharWidth = (int) layout.getAdvance();

    double width = -1;
    rows:
    for (Row row : sheet) {
      Cell cell = row.getCell(column);

      if (cell == null) {
        continue;
      }

      int colspan = 1;
      for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
        CellRangeAddress region = sheet.getMergedRegion(i);
        if (containsCell(region, row.getRowNum(), column)) {
          if (!useMergedCells) {
            // If we're not using merged cells, skip this one and move on to the next.
            continue rows;
          }
          cell = row.getCell(region.getFirstColumn());
          colspan = 1 + region.getLastColumn() - region.getFirstColumn();
        }
      }

      CellStyle style = cell.getCellStyle();
      Font font = book.getFontAt(style.getFontIndex());

      CellType cellType = cell.getCellType();

      if (cellType == CellType.STRING) {
        HSSFRichTextString rt = (HSSFRichTextString) cell.getRichStringCellValue();
        String[] lines = rt.getString().split("\\n");
        for (String line : lines) {
          String txt = line + defaultChar;
          str = new AttributedString(txt);
          copyAttributes(font, str, 0, txt.length());

          if (rt.numFormattingRuns() > 0) {
            for (int j = 0; j < line.length(); j++) {
              // FIXME используется неправильный индекс символа в строке
              int idx = rt.getFontAtIndex(j);
              if (idx != 0) {
                Font fnt = book.getFontAt(idx);
                copyAttributes(fnt, str, j, j + 1);
              }
            }
          }

          layout = new TextLayout(str.getIterator(), frc);
          if (style.getRotation() != 0) {
            /*
             * Transform the text using a scale so that it's height is increased by a multiple of the
             * leading,
             * and then rotate the text before computing the bounds. The scale results in some whitespace
             * around
             * the unrotated top and bottom of the text that normally wouldn't be present if unscaled, but
             * is added by the standard Excel autosize.
             */
            AffineTransform trans = new AffineTransform();
            trans.concatenate(
                AffineTransform.getRotateInstance(style.getRotation() * 2.0 * Math.PI / 360.0));
            trans.concatenate(AffineTransform.getScaleInstance(1, fontHeightMultiple));
            width =
                Math.max(
                    width,
                    ((layout.getOutline(trans).getBounds().getWidth() / colspan) / defaultCharWidth)
                        + cell.getCellStyle().getIndention());
          } else {
            width =
                Math.max(
                    width,
                    ((layout.getBounds().getWidth() / colspan) / defaultCharWidth)
                        + cell.getCellStyle().getIndention());
          }
        }
      } else {
        String sval = null;
        if (cellType == CellType.NUMERIC) {
          sval = getNumericCellValue(cell.getNumericCellValue(), style);
        } else if (cell.getCellType() == CellType.BOOLEAN) {
          sval = String.valueOf(cell.getBooleanCellValue());
        } else if (cellType == CellType.FORMULA) {
          FormulaEvaluator evaluator = book.getCreationHelper().createFormulaEvaluator();
          CellValue cellValue = evaluator.evaluate(cell);
          if (cellValue.getCellType() == CellType.NUMERIC) {
            sval = getNumericCellValue(cellValue.getNumberValue(), style);
          }
        }
        if (sval != null) {
          String txt = sval + defaultChar;
          str = new AttributedString(txt);
          copyAttributes(font, str, 0, txt.length());

          layout = new TextLayout(str.getIterator(), frc);
          if (style.getRotation() != 0) {
            /*
             * Transform the text using a scale so that it's height is increased by a multiple of the
             * leading,
             * and then rotate the text before computing the bounds. The scale results in some whitespace
             * around
             * the unrotated top and bottom of the text that normally wouldn't be present if unscaled, but
             * is added by the standard Excel autosize.
             */
            AffineTransform trans = new AffineTransform();
            trans.concatenate(
                AffineTransform.getRotateInstance(style.getRotation() * 2.0 * Math.PI / 360.0));
            trans.concatenate(AffineTransform.getScaleInstance(1, fontHeightMultiple));
            width =
                Math.max(
                    width,
                    ((layout.getOutline(trans).getBounds().getWidth() / colspan) / defaultCharWidth)
                        + cell.getCellStyle().getIndention());
          } else {
            width =
                Math.max(
                    width,
                    ((layout.getBounds().getWidth() / colspan) / defaultCharWidth)
                        + cell.getCellStyle().getIndention());
          }
        }
      }
    }
    if (width != -1) {
      width *= 256;
      if (width > Short.MAX_VALUE) { // width can be bigger that Short.MAX_VALUE!
        width = Short.MAX_VALUE;
      }
      sheet.setColumnWidth(column, (int) width);
    }
  }

  public static String getNumericCellValue(double value, CellStyle style) {
    String sval = null;
    String format = style.getDataFormatString().replaceAll("\"", "");
    try {
      NumberFormat fmt;
      if ("General".equals(format)) sval = "" + value;
      else {
        fmt = new DecimalFormat(format);
        sval = fmt.format(value);
      }
    } catch (Exception e) {
      sval = "" + value;
    }
    return sval;
  }
}

// CHECKSTYLE:ON
