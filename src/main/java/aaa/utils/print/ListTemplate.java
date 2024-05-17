package aaa.utils.print;

import static aaa.utils.print.DynamicTemplates.DEFAULT_STYLES;
import static aaa.utils.print.type.Types.localDateTimeType;
import static net.sf.dynamicreports.report.builder.DynamicReports.template;

import aaa.i18n.I18NResolver;
import aaa.utils.print.DynamicTemplates.Styles;
import java.time.LocalDateTime;
import lombok.experimental.UtilityClass;
import net.sf.dynamicreports.report.builder.ReportTemplateBuilder;
import net.sf.dynamicreports.report.constant.PageType;
import net.sf.dynamicreports.report.constant.SplitType;

@UtilityClass
public class ListTemplate {

  public static final String REPORT_SYSTEM_NAME = "reportSystemName";
  public static final String REPORT_CREATE_DATE = "reportCreateDate";
  public static final String PAGES = "pages_x_of_y";
  public static final String LOCALE = "locale";

  public static final ReportTemplateBuilder TEMPLATE = makeTemplate(DEFAULT_STYLES);

  public static ReportTemplateBuilder makeTemplate(Styles styles) {
    return template()
        .setPageFormat(PageType.A4)
        .setDefaultFont(styles.defaultFont)
        .setColumnStyle(styles.columnStyle)
        .setColumnTitleStyle(styles.columnTitleStyle)
        .setGroupStyle(styles.groupStyle)
        .setGroupTitleStyle(styles.groupStyle)
        .setSubtotalStyle(styles.subtotalStyle)
        .setDefaultSplitType(SplitType.PREVENT)
        .setDetailSplitType(SplitType.PREVENT);
  }

  public static MapResourceBundle commonBundle(MapResourceBundle bundle, I18NResolver i18n) {
    return bundle
        .add(
            PAGES,
            i18n.resolveCode(
                "print.pages_x_of_y", "Страница {0} из {1} {2,choice,1#страницы|1<страниц}"))
        .add(LOCALE, i18n.getLocaleCode())
        .add(REPORT_SYSTEM_NAME, i18n.resolveCode("print.systemName", "Система дистанционный офис"))
        .add(
            REPORT_CREATE_DATE,
            localDateTimeType().valueToString(LocalDateTime.now(), i18n.toJavaLocale()));
  }

  public static MapResourceBundle commonBundle(MapResourceBundle bundle) {
    return commonBundle(bundle, I18NResolver.DEFAULT_I18N);
  }
}
