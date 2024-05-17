package aaa.utils.poi;

import static aaa.utils.poi.PoiUtils.applyMergedRegions;
import static aaa.utils.poi.PoiUtils.copyStyleTo;
import static aaa.utils.poi.PoiUtils.extractRegions;
import static aaa.utils.poi.PoiUtils.shiftRegions;
import static aaa.utils.poi.PoiUtils.workbookFromBytes;
import static org.apache.commons.lang3.StringUtils.replaceChars;

import java.util.HashMap;
import java.util.Map;
import lombok.experimental.UtilityClass;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

@UtilityClass
public class MergePoiUtils {

  public static final int MAX_COLUMN_WIDTH = 255 * 255;

  public static void mergeExcelFiles(Workbook workbook, Map<String, byte[]> list) {
    list.forEach(
        (name, data) -> {
          Workbook sourceBook = workbookFromBytes(data);
          int i = 0;
          for (Sheet sourceSheet : sourceBook) {
            copySheets(
                workbook,
                workbook.createSheet(name + (sourceBook.getNumberOfSheets() > 1 ? "_" + ++i : "")),
                sourceSheet);
          }
        });
  }

  public static void copySheets(Workbook newWorkbook, Sheet newSheet, Sheet sheet) {
    int offsetRow = Math.max(newSheet.getLastRowNum(), 0);
    int maxColumnNum = 0;
    Map<Integer, CellStyle> styleMap = new HashMap<>();

    for (int i = sheet.getFirstRowNum(); i <= sheet.getLastRowNum(); i++) {
      Row srcRow = sheet.getRow(i);
      Row destRow = newSheet.createRow(i + offsetRow);
      if (srcRow != null) {
        copyRow(newWorkbook, srcRow, destRow, styleMap);
        if (srcRow.getLastCellNum() > maxColumnNum) {
          maxColumnNum = srcRow.getLastCellNum();
        }
      }
    }
    for (int i = 0; i <= maxColumnNum; i++) {
      // Math.min нужен, чтобы защититься от ошибки
      // "the maximum column width for an individual cell is 255 characters"
      newSheet.setColumnWidth(i, Math.min(MAX_COLUMN_WIDTH, sheet.getColumnWidth(i)));
    }

    PoiUtils.applyMargins(newSheet, sheet);
    applyMergedRegions(newSheet, shiftRegions(extractRegions(sheet), offsetRow, 0));
  }

  public static void copyRow(
      Workbook newWorkbook, Row srcRow, Row destRow, Map<Integer, CellStyle> styleMap) {
    destRow.setHeight(srcRow.getHeight());
    for (int j = srcRow.getFirstCellNum(); j <= srcRow.getLastCellNum(); j++) {
      Cell oldCell = srcRow.getCell(j);
      Cell newCell = destRow.getCell(j);
      if (oldCell != null) {
        if (newCell == null) {
          newCell = destRow.createCell(j);
        }
        copyCell(newWorkbook, oldCell, newCell, styleMap);
      }
    }
  }

  public static void copyCell(
      Workbook newWorkbook, Cell oldCell, Cell newCell, Map<Integer, CellStyle> styleMap) {
    int stHashCode = oldCell.getCellStyle().hashCode();
    CellStyle newCellStyle = styleMap.get(stHashCode);
    if (newCellStyle == null) {
      newCellStyle = newWorkbook.createCellStyle();
      try {
        newCellStyle.cloneStyleFrom(oldCell.getCellStyle());
      } catch (IllegalArgumentException e) {
        newCellStyle =
            copyStyleTo(oldCell.getSheet().getWorkbook(), oldCell.getCellStyle(), newWorkbook);
      }
      styleMap.put(stHashCode, newCellStyle);
    }
    newCell.setCellStyle(newCellStyle);

    CreationHelper creationHelper = newWorkbook.getCreationHelper();

    switch (oldCell.getCellType()) {
      case STRING:
        newCell.setCellValue(
            creationHelper.createRichTextString(oldCell.getRichStringCellValue().getString()));
        break;
      case FORMULA:
        // newCell.setCellFormula(oldCell.getCellFormula());
        newCell.setCellValue(
            creationHelper.createRichTextString(oldCell.getRichStringCellValue().getString()));
        break;
      case NUMERIC:
        newCell.setCellValue(oldCell.getNumericCellValue());
        break;
      case BLANK:
        newCell.setBlank();
        break;
      case BOOLEAN:
        newCell.setCellValue(oldCell.getBooleanCellValue());
        break;
      case ERROR:
        newCell.setCellErrorValue(oldCell.getErrorCellValue());
        break;

      default:
        break;
    }
  }

  public static String safeTabName(String tabName) {
    return replaceChars(tabName, "\\/*[]:?", "");
  }
}
