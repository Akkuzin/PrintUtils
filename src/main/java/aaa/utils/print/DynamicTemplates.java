package aaa.utils.print;

import static aaa.nvl.Nvl.nvl;
import static aaa.utils.print.Format.DOCX;
import static aaa.utils.print.Format.HTML;
import static aaa.utils.print.Format.PDF;
import static aaa.utils.print.Format.RTF;
import static aaa.utils.print.Format.TXT;
import static aaa.utils.print.Format.isExcel;
import static aaa.utils.print.Format.isOOXML;
import static aaa.utils.print.ListTemplate.REPORT_CREATE_DATE;
import static aaa.utils.print.ListTemplate.REPORT_SYSTEM_NAME;
import static java.util.Arrays.asList;
import static java.util.Collections.emptyMap;
import static net.sf.dynamicreports.report.builder.component.Components.horizontalList;
import static net.sf.dynamicreports.report.builder.component.Components.text;
import static net.sf.dynamicreports.report.builder.component.Components.verticalList;
import static net.sf.dynamicreports.report.builder.expression.Expressions.jasperSyntax;
import static net.sf.dynamicreports.report.builder.style.Styles.font;
import static net.sf.dynamicreports.report.builder.style.Styles.pen;
import static net.sf.dynamicreports.report.builder.style.Styles.penThin;
import static net.sf.dynamicreports.report.builder.style.Styles.style;
import static net.sf.dynamicreports.report.constant.HorizontalTextAlignment.CENTER;
import static net.sf.dynamicreports.report.constant.HorizontalTextAlignment.LEFT;
import static net.sf.dynamicreports.report.constant.VerticalTextAlignment.MIDDLE;
import static net.sf.dynamicreports.report.constant.VerticalTextAlignment.TOP;
import static net.sf.jasperreports.engine.JasperFillManager.fillReport;

import aaa.basis.text.StringFunc;
import com.google.common.collect.ImmutableMap;

import java.awt.*;
import java.io.ByteArrayOutputStream;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.util.HashMap;
import java.util.Locale;
import java.util.Map;
import java.util.function.Supplier;
import java.util.stream.Stream;
import lombok.AccessLevel;
import lombok.Getter;
import lombok.SneakyThrows;
import lombok.experimental.FieldDefaults;
import lombok.experimental.UtilityClass;
import net.sf.dynamicreports.report.builder.column.Columns;
import net.sf.dynamicreports.report.builder.column.TextColumnBuilder;
import net.sf.dynamicreports.report.builder.component.ComponentBuilder;
import net.sf.dynamicreports.report.builder.component.HorizontalListBuilder;
import net.sf.dynamicreports.report.builder.component.HorizontalListCellBuilder;
import net.sf.dynamicreports.report.builder.component.TextFieldBuilder;
import net.sf.dynamicreports.report.builder.component.VerticalListBuilder;
import net.sf.dynamicreports.report.builder.component.VerticalListCellBuilder;
import net.sf.dynamicreports.report.builder.expression.AddExpression;
import net.sf.dynamicreports.report.builder.expression.JasperExpression;
import net.sf.dynamicreports.report.builder.style.FontBuilder;
import net.sf.dynamicreports.report.builder.style.PenBuilder;
import net.sf.dynamicreports.report.builder.style.StyleBuilder;
import net.sf.dynamicreports.report.constant.LineStyle;
import net.sf.dynamicreports.report.definition.datatype.DRIDataType;
import net.sf.dynamicreports.report.definition.expression.DRIExpression;
import net.sf.jasperreports.engine.JRAbstractExporter;
import net.sf.jasperreports.engine.JRDataSource;
import net.sf.jasperreports.engine.JRException;
import net.sf.jasperreports.engine.JRParameter;
import net.sf.jasperreports.engine.JasperPrint;
import net.sf.jasperreports.engine.export.HtmlExporter;
import net.sf.jasperreports.engine.export.JRPdfExporter;
import net.sf.jasperreports.engine.export.JRRtfExporter;
import net.sf.jasperreports.engine.export.JRTextExporter;
import net.sf.jasperreports.engine.export.JRXlsAbstractExporter;
import net.sf.jasperreports.engine.export.JRXlsExporter;
import net.sf.jasperreports.engine.export.ooxml.JRDocxExporter;
import net.sf.jasperreports.engine.export.ooxml.JRXlsxExporter;
import net.sf.jasperreports.export.SimpleDocxExporterConfiguration;
import net.sf.jasperreports.export.SimpleExporterInput;
import net.sf.jasperreports.export.SimpleHtmlExporterConfiguration;
import net.sf.jasperreports.export.SimpleHtmlExporterOutput;
import net.sf.jasperreports.export.SimpleOutputStreamExporterOutput;
import net.sf.jasperreports.export.SimplePdfExporterConfiguration;
import net.sf.jasperreports.export.SimpleRtfExporterConfiguration;
import net.sf.jasperreports.export.SimpleTextReportConfiguration;
import net.sf.jasperreports.export.SimpleWriterExporterOutput;
import net.sf.jasperreports.export.SimpleXlsReportConfiguration;
import net.sf.jasperreports.export.SimpleXlsxReportConfiguration;
import org.apache.commons.lang3.ArrayUtils;

@UtilityClass
public class DynamicTemplates {

  static final Integer A4_TEXT_WIDTH = 120;
  static final Integer A4_TEXT_HEIGHT = 170;

  public static final int DEFAULT_MARGIN = 20;

  public static final String SANS = "DejaVu Sans Condensed";

  public static final int FONT_3 = 3;
  public static final int FONT_4 = 4;
  public static final int FONT_5 = 5;
  public static final int FONT_6 = 6;
  public static final int FONT_7 = 7;
  public static final int FONT_8 = 8;
  public static final int FONT_9 = 9;
  public static final int FONT_10 = 10;
  public static final int FONT_11 = 11;
  public static final int FONT_12 = 12;
  public static final int FONT_14 = 14;

  public static final PenBuilder ZERO_LINE = pen(0f, LineStyle.SOLID);

  public static final FontBuilder DEFAULT_FONT = font().setFontName(SANS).setFontSize(FONT_9);

  public static final StyleBuilder HIDE_STYLE = style().setFontSize(0).setPadding(0);

  public static final Styles DEFAULT_STYLES = new Styles(FONT_9);

  @Getter
  @FieldDefaults(level = AccessLevel.PUBLIC, makeFinal = true)
  public class Styles {

    int baseFontSize;

    public FontBuilder makeFont(int size) {
      return font().setFontName(SANS).setFontSize(size);
    }

    public <T extends StyleBuilder> T makeHeaderDecoration(T style) {
      style
          .setFontSize(baseFontSize - 1)
          .bold()
          .setBorder(penThin())
          .setBackgroundColor(new Color(240, 240, 240));
      return style;
    }

    FontBuilder defaultFont;
    StyleBuilder rooStyle;
    StyleBuilder bold;
    StyleBuilder italic;
    StyleBuilder boldCentered;
    StyleBuilder noteStyle;

    StyleBuilder titleStyle;
    StyleBuilder columnStyle;
    StyleBuilder columnTitleStyle;
    StyleBuilder groupStyle;
    StyleBuilder subtotalStyle;

    public Styles(int baseFontSize) {
      this.baseFontSize = baseFontSize;

      defaultFont = makeFont(baseFontSize);
      rooStyle = style().setFont(defaultFont).setPadding(2);
      bold = style(rooStyle).bold();
      italic = style(rooStyle).italic();
      boldCentered = style(bold).setTextAlignment(CENTER, MIDDLE);
      noteStyle = style(rooStyle).setFontSize(baseFontSize - 3);

      titleStyle = style(rooStyle).bold().setHorizontalTextAlignment(CENTER);
      columnStyle =
          style(rooStyle)
              .setFont(font().setFontName(SANS).setFontSize(baseFontSize))
              .setVerticalTextAlignment(TOP)
              .setBorder(penThin());
      columnTitleStyle =
          makeHeaderDecoration(style(columnStyle).setHorizontalTextAlignment(CENTER));
      groupStyle = style(bold).setHorizontalTextAlignment(LEFT);
      subtotalStyle = style(bold);
    }
  }

  public static final YesNoFormatter YES_NO_FORMATTER = new YesNoFormatter();

  public static final String FIELD_RESOURCE_PREFIX = "field.";

  static final String LAST_PAGE_PARAMETER_NAME = "LastPageNumber";
  static final String DETAILS_LAST_PAGE_PARAMETER_NAME = "DetailsLastPageNumber";
  static final String AFTER_TOTAL_PARAMETER_NAME = "AfterTotal";

  public static final JasperExpression<Boolean> SET_LAST_PAGE_EXPRESSION =
      jasperSyntax(
          "new Boolean($P{REPORT_PARAMETERS_MAP}.put(\""
              + LAST_PAGE_PARAMETER_NAME
              + "\",$V{PAGE_NUMBER}).equals(\"dummyPrintWhen\"))",
          Boolean.class);
  public static final JasperExpression<Boolean> SET_CURRENT_DETAIL_PAGE_EXPRESSION =
      jasperSyntax(
          "new Boolean($P{REPORT_PARAMETERS_MAP}.put(\""
              + DETAILS_LAST_PAGE_PARAMETER_NAME
              + "\",$V{PAGE_NUMBER}).equals(\"dummyPrintWhen\"))",
          Boolean.class);
  public static final JasperExpression<Boolean> SET_AFTER_TOTAL_EXPRESSION =
      jasperSyntax(
          "new Boolean($P{REPORT_PARAMETERS_MAP}.put(\""
              + AFTER_TOTAL_PARAMETER_NAME
              + "\", 1).equals(\"dummyPrintWhen\"))",
          Boolean.class);

  public static final JasperExpression<Boolean> IS_LAST_PAGE_EXPRESSION =
      jasperSyntax(
          "$P{REPORT_PARAMETERS_MAP}.get(\"" + LAST_PAGE_PARAMETER_NAME + "\") != null ",
          Boolean.class);
  public static final JasperExpression<Boolean> IS_NOT_LAST_PAGE_EXPRESSION =
      jasperSyntax(
          "$P{REPORT_PARAMETERS_MAP}.get(\"" + LAST_PAGE_PARAMETER_NAME + "\") == null ",
          Boolean.class);

  public static final JasperExpression<Boolean> IS_NOT_DETAILS_PAGE_EXPRESSION =
      jasperSyntax(
          "$P{REPORT_PARAMETERS_MAP}.get(\""
              + DETAILS_LAST_PAGE_PARAMETER_NAME
              + "\") < $V{PAGE_NUMBER}",
          Boolean.class);
  public static final JasperExpression<Boolean> IS_DETAILS_PAGE_EXPRESSION =
      jasperSyntax(
          "$P{REPORT_PARAMETERS_MAP}.get(\""
              + DETAILS_LAST_PAGE_PARAMETER_NAME
              + "\") == $V{PAGE_NUMBER}",
          Boolean.class);

  public static final JasperExpression<Boolean> IS_AFTER_TOTAL_EXPRESSION =
      jasperSyntax(
          "$P{REPORT_PARAMETERS_MAP}.get(\"" + AFTER_TOTAL_PARAMETER_NAME + "\") != null",
          Boolean.class);
  public static final JasperExpression<Boolean> IS_NOT_AFTER_TOTAL_EXPRESSION =
      jasperSyntax(
          "$P{REPORT_PARAMETERS_MAP}.get(\"" + AFTER_TOTAL_PARAMETER_NAME + "\") == null",
          Boolean.class);

  public static DRIExpression<String> resource(String name) {
    return jasperSyntax(resourceExpression(name), String.class);
  }

  public static String resourceExpression(String name) {
    return "$R{" + name + "}";
  }

  public static String fieldExpression(String name) {
    return "$F{" + name + "}";
  }

  public static String variableExpression(String name) {
    return "$V{" + name + "}";
  }

  public static String parameterExpression(String name) {
    return "$P{" + name + "}";
  }

  public static <T> TextColumnBuilder<T> column(
      String fieldName, DRIDataType<? super T, T> dataType) {
    return nestedColumn(null, fieldName, dataType);
  }

  public static <T> TextColumnBuilder<T> columnPlaceholder(
      String fieldName, Class<T> clazz, Map<String, Integer> widths) {
    return Columns.column(jasperSyntax("null", clazz))
        .setWidth(widths.get(fieldName))
        .setTitle(columnTitle(fieldName));
  }

  public static <T> TextColumnBuilder<T> column(
      String fieldName, DRIDataType<? super T, T> dataType, Map<String, Integer> widths) {
    return nestedColumn(null, fieldName, dataType, widths);
  }

  public static <T> TextColumnBuilder<T> nestedColumn(
      String prefix, String fieldName, DRIDataType<? super T, T> dataType) {
    return Columns.column(StringFunc.concatWithDelim(prefix, ".", fieldName), dataType)
        .setTitle(columnTitle(fieldName));
  }

  public static <T> TextColumnBuilder<T> nestedColumn(
      String prefix,
      String fieldName,
      DRIDataType<? super T, T> dataType,
      Map<String, Integer> widths) {
    return nestedColumn(prefix, fieldName, dataType).setWidth(widths.get(fieldName));
  }

  public static TextColumnBuilder<String> columnI18n(
      String fieldName, Map<String, Integer> widths) {
    return Columns.column(
            jasperSyntax(
                fieldExpression(fieldName)
                    + ".get("
                    + resourceExpression(ListTemplate.LOCALE)
                    + ")",
                String.class))
        .setTitle(columnTitle(fieldName))
        .setWidth(widths.get(fieldName));
  }

  public static String nameEntityFieldExpression(String namesParameter, String fieldName) {
    return parameterExpression(namesParameter)
        + ".get("
        + fieldExpression(fieldName)
        + ".getCode())";
  }

  public static TextColumnBuilder<String> columnNamedEntity(
      String fieldName, String namesParameter, Map<String, Integer> widths) {
    return Columns.column(
            jasperSyntax(nameEntityFieldExpression(namesParameter, fieldName), String.class))
        .setTitle(columnTitle(fieldName))
        .setWidth(widths.get(fieldName));
  }

  public static TextFieldBuilder<String> makeSystemInfoFooter(Styles styles) {
    return text(jasperSyntax(
            resourceExpression(REPORT_SYSTEM_NAME)
                + "+\" (\"+"
                + resourceExpression(REPORT_CREATE_DATE)
                + "+\")\"",
            String.class))
        .setStyle(styles.noteStyle)
        .setHorizontalTextAlignment(LEFT);
  }

  public static DRIExpression<String> columnTitle(String fieldName) {
    return resource(FIELD_RESOURCE_PREFIX + fieldName);
  }

  public static AddExpression sumExpression(DRIExpression<? extends Number>... expressions) {
    return new AddExpression(expressions);
  }

  public static HorizontalListBuilder hList() {
    return horizontalList();
  }

  public static HorizontalListBuilder hList(ComponentBuilder<?, ?>... components) {
    return horizontalList(components);
  }

  public static HorizontalListBuilder hList(HorizontalListCellBuilder... cells) {
    return horizontalList(cells);
  }

  public static VerticalListBuilder vList() {
    return verticalList();
  }

  public static VerticalListBuilder vList(ComponentBuilder<?, ?>... components) {
    return verticalList(components);
  }

  public static VerticalListBuilder vList(VerticalListCellBuilder... cells) {
    return verticalList(cells);
  }

  public static class ParamsBuilder {

    Map<String, Object> params;

    public ParamsBuilder(Map<String, Object> params) {
      this.params = new HashMap<>(params);
    }

    public ParamsBuilder add(String name, Object value) {
      params.put(name, value);
      return this;
    }

    public Map<String, Object> get() {
      return this.params;
    }

    public static <T> Map<String, T> concatMaps(Map<String, T>... maps) {
      return Stream.of(maps).collect(HashMap::new, Map::putAll, Map::putAll);
    }

    public static ParamsBuilder from(MapResourceBundle bundle) {
      return from(bundle, new java.util.Locale("ru"));
    }

    public static ParamsBuilder from(MapResourceBundle bundle, Locale locale) {
      return new ParamsBuilder(
          ImmutableMap.of(
              JRParameter.REPORT_RESOURCE_BUNDLE,
              bundle,
              JRParameter.REPORT_LOCALE,
              locale,
              JRParameter.REPORT_FORMAT_FACTORY,
              new FormatFactory()));
    }

    public ParamsBuilder on(Report report) {
      return new ParamsBuilder(concatMaps(report.parameters, nvl(this.params, emptyMap())));
    }
  }

  public static byte[] outputReport(
      Supplier<Report> report, ParamsBuilder params, JRDataSource dataSource) {
    return outputReport(report.get(), null, params, dataSource);
  }

  @SneakyThrows
  public static byte[] outputReport(
      Report report,
      Format format,
      DynamicTemplates.ParamsBuilder params,
      JRDataSource dataSource) {
    return outputReport(
        nvl(format, report.format), fillReport(report.report, params.on(report).get(), dataSource));
  }

  public static byte[] outputReport(Format format, JasperPrint... print) {
    ByteArrayOutputStream out = new ByteArrayOutputStream();
    outputReport(format, out, print);
    return out.toByteArray();
  }

  public JRPdfExporter pdfExporter() {
    JRPdfExporter exporter = new JRPdfExporter();
    SimplePdfExporterConfiguration config = new SimplePdfExporterConfiguration();
    config.setCompressed(true);
    exporter.setConfiguration(config);
    return exporter;
  }

  public static JRXlsAbstractExporter excelExporter(Format format) {
    boolean isXssf = isOOXML(format);
    JRXlsAbstractExporter exporter = isXssf ? new JRXlsxExporter() : new JRXlsExporter();
    SimpleXlsReportConfiguration config =
        isXssf ? new SimpleXlsxReportConfiguration() : new SimpleXlsReportConfiguration();
    // config.setFontSizeFixEnabled(true);
    if (format.defaultPaging) {
      config.setForcePageBreaks(true);
      config.setRemoveEmptySpaceBetweenRows(true);
      config.setCollapseRowSpan(true);
    } else {
      config.setRemoveEmptySpaceBetweenRows(false);
      config.setCollapseRowSpan(false);
    }
    config.setDetectCellType(true);
    config.setIgnoreCellBorder(false);
    exporter.setConfiguration(config);
    return exporter;
  }

  public static JRTextExporter textExporter() {
    JRTextExporter exporter = new JRTextExporter();
    SimpleTextReportConfiguration config = new SimpleTextReportConfiguration();
    config.setPageWidthInChars(A4_TEXT_WIDTH);
    config.setPageHeightInChars(A4_TEXT_HEIGHT);
    exporter.setConfiguration(config);
    return exporter;
  }

  public static HtmlExporter htmlExporter() {
    HtmlExporter exporter = new HtmlExporter();
    SimpleHtmlExporterConfiguration config = new SimpleHtmlExporterConfiguration();
    exporter.setConfiguration(config);
    return exporter;
  }

  public static JRDocxExporter docxExporter() {
    JRDocxExporter exporter = new JRDocxExporter();
    SimpleDocxExporterConfiguration config = new SimpleDocxExporterConfiguration();
    config.setEmbedFonts(true);
    exporter.setConfiguration(config);
    return exporter;
  }

  public static JRRtfExporter rtfExporter() {
    JRRtfExporter exporter = new JRRtfExporter();
    exporter.setConfiguration(new SimpleRtfExporterConfiguration());
    return exporter;
  }

  @SneakyThrows
  public static void outputReport(Format format, OutputStream out, JasperPrint... print) {
    if (format == PDF) {
      JRPdfExporter exporter = pdfExporter();
      exporter.setExporterOutput(new SimpleOutputStreamExporterOutput(out));
      export(exporter, print);
    } else if (isExcel(format)) {
      if (!format.defaultPaging && ArrayUtils.isNotEmpty(print)) {
        Stream.of(print)
            .forEach(
                singlePrint ->
                    singlePrint.setProperty(JRXlsxExporter.PROPERTY_AUTO_FIT_ROW, "true"));
      }
      JRXlsAbstractExporter exporter = excelExporter(format);
      exporter.setExporterOutput(new SimpleOutputStreamExporterOutput(out));
      export(exporter, print);
    } else if (format == HTML) {
      HtmlExporter exporter = htmlExporter();
      exporter.setExporterOutput(new SimpleHtmlExporterOutput(out, StandardCharsets.UTF_8.name()));
      export(exporter, print);
    } else if (format == DOCX) {
      JRDocxExporter exporter = docxExporter();
      exporter.setExporterOutput(new SimpleOutputStreamExporterOutput(out));
      export(exporter, print);
    } else if (format == RTF) {
      JRRtfExporter exporter = rtfExporter();
      exporter.setExporterOutput(
          new SimpleWriterExporterOutput(out, StandardCharsets.UTF_8.name()));
      export(exporter, print);
    } else if (format == TXT) {
      JRTextExporter exporter = textExporter();
      exporter.setExporterOutput(
          new SimpleWriterExporterOutput(out, StandardCharsets.UTF_8.name()));
      export(exporter, print);
    }
  }

  public static void export(JRAbstractExporter exporter, JasperPrint... print) throws JRException {
    exporter.setExporterInput(SimpleExporterInput.getInstance(asList(print)));
    exporter.exportReport();
  }
}
