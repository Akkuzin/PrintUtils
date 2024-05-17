package aaa.utils.print;

import net.sf.jasperreports.engine.export.JRExporterGridCell;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.ss.util.CellRangeAddress;

@Deprecated
public class JRXlsExporterFix extends net.sf.jasperreports.engine.export.JRXlsExporter {

  /** Оптимизация производительности addMergedRegion => addMergedRegionUnsafe */
  protected void createMergeRegion(
      JRExporterGridCell gridCell, int colIndex, int rowIndex, HSSFCellStyle cellStyle) {
    boolean isCollapseRowSpan = getCurrentItemConfiguration().isCollapseRowSpan();
    int rowSpan = isCollapseRowSpan ? 1 : gridCell.getRowSpan();
    if (gridCell.getColSpan() > 1 || rowSpan > 1) {
      sheet.addMergedRegionUnsafe(
          new CellRangeAddress(
              rowIndex, rowIndex + rowSpan - 1, colIndex, colIndex + gridCell.getColSpan() - 1));

      for (int i = 0; i < rowSpan; i++) {
        HSSFRow spanRow = sheet.getRow(rowIndex + i);
        if (spanRow == null) {
          spanRow = sheet.createRow(rowIndex + i);
        }
        for (int j = 0; j < gridCell.getColSpan(); j++) {
          HSSFCell spanCell = spanRow.getCell(colIndex + j);
          if (spanCell == null) {
            spanCell = spanRow.createCell(colIndex + j);
          }
          spanCell.setCellStyle(cellStyle);
        }
      }
    }
  }
}
