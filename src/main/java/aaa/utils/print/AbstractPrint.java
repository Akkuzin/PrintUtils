package aaa.utils.print;

import static aaa.utils.print.DynamicTemplates.outputReport;
import static aaa.utils.print.Format.ALL_EXCEL;
import static aaa.utils.print.Format.ALL_WORD;
import static aaa.utils.print.Format.PDF;
import static aaa.utils.print.PrintUtils.makeReport;
import static aaa.utils.print.PrintUtils.mapConcat;
import static java.util.Collections.emptyMap;
import static java.util.Optional.ofNullable;

import aaa.i18n.I18NResolver;
import aaa.utils.print.DynamicTemplates.ParamsBuilder;
import com.google.common.collect.Streams;
import java.util.Map;
import java.util.stream.Stream;
import lombok.Data;
import lombok.Getter;
import lombok.RequiredArgsConstructor;
import lombok.SneakyThrows;
import net.sf.dynamicreports.jasper.builder.JasperReportBuilder;
import net.sf.jasperreports.engine.JRDataSource;

public abstract class AbstractPrint<P extends IPrintParameters, RT extends ReportType>
    implements IPrint<P> {

  @Getter(lazy = true)
  private final Map<RT, Report> reports = makeReports();

  protected Map<RT, Report> makeReports() {
    return availableReportType()
        .map(type -> makeReport(type, this::makeReportTemplate))
        .reduce(mapConcat())
        .orElse(emptyMap());
  }

  protected Stream<RT> availableReportType() {
    return Streams.concat(Stream.of(PDF), ALL_WORD.stream(), ALL_EXCEL.stream())
        .map(format -> makeReportType((P) SimplePrintParameters.of(format)));
  }

  protected abstract JasperReportBuilder makeReportTemplate(RT reportType);

  protected abstract JRDataSource makeDataSource(P params);

  protected MapResourceBundle makeBundle(P params) {
    MapResourceBundle bundle =
        ListTemplate.commonBundle(new MapResourceBundle(), params.getI18nSafe());
    fillBundle(params, bundle);
    return bundle;
  }

  protected void fillBundle(P params, MapResourceBundle bundle) {}

  @Data
  @RequiredArgsConstructor(staticName = "of")
  private static class SimplePrintParameters implements IPrintParameters {
    final Format format;
    I18NResolver i18n;
  }

  protected RT makeReportType(P params) {
    return (RT)
        ReportType.of(ofNullable(params).map(IPrintParameters::getFormatSafe).orElse(Format.PDF));
  }

  @SneakyThrows
  public byte[] generateBlank(P params) {
    return outputReport(
        () -> getReports().get(makeReportType(params)),
        ParamsBuilder.from(makeBundle(params)),
        makeDataSource(params));
  }
}
