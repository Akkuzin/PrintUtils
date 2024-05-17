package aaa.utils.print.expression;

import static net.sf.jasperreports.engine.JRParameter.REPORT_FORMAT_FACTORY;
import static net.sf.jasperreports.engine.JRParameter.REPORT_LOCALE;

import java.util.HashMap;
import java.util.List;
import java.util.Map;
import lombok.experimental.FieldDefaults;
import net.sf.dynamicreports.report.base.expression.AbstractSimpleExpression;
import net.sf.dynamicreports.report.definition.ReportParameters;

@FieldDefaults(makeFinal = true)
public class SubreportParameterExpression extends AbstractSimpleExpression<Map<String, Object>> {

  List<String> parameters;

  protected SubreportParameterExpression(List<String> parameters) {
    this.parameters = parameters;
  }

  public static SubreportParameterExpression of(List<String> parameters) {
    return new SubreportParameterExpression(parameters);
  }

  public Map<String, Object> evaluate(ReportParameters reportParameters) {
    Map<String, Object> map = new HashMap<>();
    for (String name : parameters) {
      map.put(name, reportParameters.getParameterValue(name));
    }
    map.put(REPORT_LOCALE, reportParameters.getParameterValue(REPORT_LOCALE));
    map.put(REPORT_FORMAT_FACTORY, reportParameters.getParameterValue(REPORT_FORMAT_FACTORY));
    return map;
  }
}
