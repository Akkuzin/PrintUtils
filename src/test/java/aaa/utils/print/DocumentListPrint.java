package aaa.utils.print;

import static aaa.utils.print.DynamicTemplates.DEFAULT_MARGIN;
import static aaa.utils.print.DynamicTemplates.column;
import static aaa.utils.print.Format.PDF;
import static java.util.Arrays.asList;
import static net.sf.dynamicreports.report.builder.DynamicReports.margin;
import static net.sf.dynamicreports.report.builder.DynamicReports.report;
import static net.sf.dynamicreports.report.builder.datatype.DataTypes.longType;
import static net.sf.dynamicreports.report.builder.datatype.DataTypes.stringType;

import aaa.i18n.I18NResolver;
import aaa.utils.print.DocumentListPrint.Document.Fields;
import aaa.utils.print.DocumentListPrint.Parameters;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.List;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Builder.Default;
import lombok.Data;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import lombok.SneakyThrows;
import lombok.experimental.FieldNameConstants;
import net.sf.dynamicreports.jasper.builder.JasperReportBuilder;
import net.sf.dynamicreports.report.constant.PageOrientation;
import net.sf.dynamicreports.report.constant.PageType;
import net.sf.jasperreports.engine.JRDataSource;
import net.sf.jasperreports.engine.data.JRBeanCollectionDataSource;
import org.junit.jupiter.api.Test;

public class DocumentListPrint extends AbstractPrint<Parameters, ReportType> {

  protected JasperReportBuilder makeReportTemplate(ReportType type) {
    JasperReportBuilder report =
        report()
            .setTemplate(ListTemplate.TEMPLATE)
            .setPageFormat(PageType.A4, PageOrientation.PORTRAIT)
            .setPageMargin(margin(DEFAULT_MARGIN).setLeft(60).setRight(30))
            .setSummaryWithPageHeaderAndFooter(true);

    report.columns(column(Fields.id, longType()), column(Fields.name, stringType()));

    return report;
  }

  @Override
  protected JRDataSource makeDataSource(Parameters params) {
    return new JRBeanCollectionDataSource(params.documents);
  }

  @Override
  protected void fillBundle(Parameters params, MapResourceBundle bundle) {
    I18NResolver i18n = params.getI18nSafe();
    bundle.fieldTitle(Fields.id, "ID");
    bundle.fieldTitle(Fields.name, "Наименование");
  }

  @Builder
  @Getter
  @Setter
  public static class Parameters implements IPrintParameters {
    List<Document> documents;
    @Default Format format = PDF;
    I18NResolver i18n;
  }

  @SneakyThrows
  @Test
  public void testPrint() {
    DocumentListPrint print = new DocumentListPrint();
    Parameters params =
        Parameters.builder()
            .documents(
                asList(
                    new Document(1L, "Doc 1"),
                    new Document(2L, "Doc 2"),
                    new Document(3L, "Doc 3")))
            .build();
    Files.write(
        Paths.get(System.getProperty("user.dir") + "/target/test_list_print.pdf"),
        print.generateBlank(params));
  }

  @Data
  @AllArgsConstructor
  @NoArgsConstructor
  @FieldNameConstants
  public static class Document {

    Long id;
    String name;
  }
}
