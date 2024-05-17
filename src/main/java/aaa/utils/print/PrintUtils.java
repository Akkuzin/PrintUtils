package aaa.utils.print;

import static aaa.utils.poi.PoiUtils.applyMargins;
import static aaa.utils.poi.PoiUtils.cleanRowBreaks;
import static aaa.utils.poi.PoiUtils.flushWorkBook;
import static aaa.utils.poi.PoiUtils.simpleCollectToSingle;
import static aaa.utils.print.Format.PDF;
import static aaa.utils.print.Format.isExcel;
import static aaa.utils.print.Format.isOOXML;
import static java.util.stream.Collectors.toList;
import static net.sf.dynamicreports.jasper.base.JasperScriptletManager.USE_THREAD_SAFE_SCRIPLET_MANAGER_PROPERTY_KEY;

import aaa.utils.pdf.PdfUtils;
import aaa.utils.poi.MergePoiUtils;
import aaa.utils.poi.PoiUtils;
import com.google.common.collect.ImmutableMap;
import java.io.ByteArrayOutputStream;
import java.lang.reflect.Field;
import java.util.Collection;
import java.util.Map;
import java.util.Properties;
import java.util.concurrent.atomic.AtomicReference;
import java.util.function.BinaryOperator;
import java.util.function.Consumer;
import java.util.function.Function;
import java.util.function.Supplier;
import lombok.SneakyThrows;
import lombok.experimental.UtilityClass;
import net.sf.dynamicreports.jasper.builder.JasperReportBuilder;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.pdfbox.io.RandomAccessReadBuffer;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

@UtilityClass
public class PrintUtils {

  public static int sumColumnsWidths(Map<String, Integer> widths) {
    return widths.values().stream().mapToInt(Integer::intValue).sum();
  }

  public static <RT extends ReportType> BinaryOperator<Map<RT, Report>> mapConcat() {
    return (map1, map2) -> ImmutableMap.<RT, Report>builder().putAll(map1).putAll(map2).build();
  }

  public static JasperReportBuilder prepareReport(JasperReportBuilder report) {
    Properties properties = new Properties();
    properties.setProperty(USE_THREAD_SAFE_SCRIPLET_MANAGER_PROPERTY_KEY, "true");
    report.setProperties(properties);
    return report;
  }

  @SneakyThrows
  public static <T extends ReportType> Map<T, Report> makeReport(
      T reportType, Function<T, JasperReportBuilder> reportSupplier) {
    return Map.of(
        reportType,
        Report.of(prepareReport(reportSupplier.apply(reportType)), reportType.getFormat()));
  }

  public static byte[] mergeTabsToFormat(
      Map<String, Supplier<byte[]>> data, Format format, Consumer<Workbook> postProcessor) {
    if (format == PDF) {
      ByteArrayOutputStream os = new ByteArrayOutputStream();
      PdfUtils.doMerge(
          os,
          data.values().stream()
              .map(Supplier::get)
              .filter(ArrayUtils::isNotEmpty)
              .map(RandomAccessReadBuffer::new)
              .collect(toList()));
      return os.toByteArray();
    } else if (isExcel(format)) {
      Workbook output = isOOXML(format) ? new XSSFWorkbook() : new HSSFWorkbook();
      data.forEach(
          (name, byteData) -> {
            byte[] workbookData = byteData.get();
            if (ArrayUtils.isNotEmpty(workbookData)) {
              Workbook source = PoiUtils.workbookFromBytes(workbookData);
              Sheet sheet = output.createSheet();
              output.setSheetName(output.getSheetIndex(sheet), MergePoiUtils.safeTabName(name));
              applyMargins(sheet, source.iterator().next());
              simpleCollectToSingle(source, sheet, output);
            }
          });
      cleanRowBreaks(output);
      if (postProcessor != null) {
        postProcessor.accept(output);
      }
      return flushWorkBook(output);
    }
    return null;
  }

  public static byte[] mergeTabsToFormat(Map<String, Supplier<byte[]>> data, Format format) {
    return mergeTabsToFormat(data, format, null);
  }

  public static void clearPrints(Collection<IPrint> prints) {
    prints.forEach(
        print -> {
          try {
            Field field = print.getClass().getDeclaredField("reports");
            field.setAccessible(true);
            ((AtomicReference) field.get(print)).set(null);
          } catch (NoSuchFieldException | IllegalAccessException e) {
          }
        });
  }
}
