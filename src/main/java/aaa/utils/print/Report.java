package aaa.utils.print;

import java.util.Collections;
import java.util.Map;
import lombok.AccessLevel;
import lombok.Builder;
import lombok.experimental.FieldDefaults;
import net.sf.dynamicreports.jasper.builder.JasperReportBuilder;
import net.sf.dynamicreports.report.exception.DRException;
import net.sf.jasperreports.engine.JasperReport;

@Builder
@FieldDefaults(level = AccessLevel.PUBLIC, makeFinal = true)
public class Report {

  Map<String, Object> parameters;
  JasperReport report;
  Format format;

  public static Report of(JasperReportBuilder reportBuilder, Format format) throws DRException {
    return Report.builder()
        .report(reportBuilder.toJasperReport())
        .parameters(Collections.unmodifiableMap(reportBuilder.getJasperParameters()))
        .format(format)
        .build();
  }
}
