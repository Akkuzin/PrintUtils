package aaa.utils.poi;


import aaa.utils.poi.formula.BaseFormula;
import aaa.utils.poi.formula.ConstantFormula;
import aaa.utils.poi.formula.DivideFormula;
import aaa.utils.poi.formula.EqualFormula;
import aaa.utils.poi.formula.IBaseFormula;
import aaa.utils.poi.formula.IfFormula;
import aaa.utils.poi.formula.MultiplyFormula;
import java.util.ArrayList;
import java.util.List;
import lombok.experimental.UtilityClass;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;

@UtilityClass
public class FormulaPoiUtils {

  private static final char ALPHABET_BASE = 'A';

  private static String getLetterByIndex(long index) {
    return String.valueOf((char) (index % 26 + ALPHABET_BASE));
  }

  /** Returns symbolic name of the Excel column */
  public static String columnIndexToName(long index) {
    List<String> list = new ArrayList<>();
    long first = 0L;
    for (long i = index; i > 0; ) {
      list.add(0, getLetterByIndex(i % 26 - first));
      i /= 26;
      first = 1L;
    }
    return String.join("", list);
  }

  public static String rowIndexToName(long index) {
    return String.valueOf(index + 1);
  }

  /** Returns symbolic name of the Excel cell */
  public static String getCellAddress(long rowIndex, long columnIndex) {
    return getCellAddress(rowIndexToName(rowIndex), columnIndexToName(columnIndex));
  }

  public static String getCellAddress(long rowIndex, String columnIndex) {
    return getCellAddress(rowIndexToName(rowIndex), columnIndex);
  }

  public static String getCellAddress(String rowIndex, long columnIndex) {
    return getCellAddress(rowIndex, columnIndexToName(columnIndex));
  }

  public static String getCellAddress(String rowIndex, String columnIndex) {
    return columnIndex + rowIndex;
  }

  public static BaseFormula makeRatioFormula(IBaseFormula part, IBaseFormula whole) {
    return new IfFormula(
        new EqualFormula(whole, new ConstantFormula(0)),
        new ConstantFormula(0),
        new MultiplyFormula(new DivideFormula(part, whole), new ConstantFormula(100)));
  }

  public static void recalculateAllFormulae(HSSFWorkbook workBook) {
    FormulaEvaluator evaluator = new HSSFFormulaEvaluator(workBook);
    for (int sheetNum = 0; sheetNum < workBook.getNumberOfSheets(); sheetNum++) {
      for (Row row : workBook.getSheetAt(sheetNum)) {
        for (Cell cell : row) {
          if (cell.getCellType() == CellType.FORMULA) {
            evaluator.evaluateFormulaCell(cell);
          }
        }
      }
    }
  }
}
