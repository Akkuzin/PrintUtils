package aaa.utils.print;

import static java.lang.Boolean.TRUE;
import static net.sf.jasperreports.engine.JRParameter.REPORT_RESOURCE_BUNDLE;

import net.sf.dynamicreports.report.base.expression.AbstractValueFormatter;
import net.sf.dynamicreports.report.definition.ReportParameters;

public class YesNoFormatter extends AbstractValueFormatter<String, Boolean> {

  public static final String TRUE_CODE = "constants.true";
  public static final String FALSE_CODE = "constants.false";
  public static final String NULL_CODE = "constants.null";

  @Override
  public String format(Boolean value, ReportParameters reportParameters) {
    return reportParameters
        .<MapResourceBundle>getParameterValue(REPORT_RESOURCE_BUNDLE)
        .getStringSafe(value == null ? NULL_CODE : TRUE.equals(value) ? TRUE_CODE : FALSE_CODE);
  }
}
