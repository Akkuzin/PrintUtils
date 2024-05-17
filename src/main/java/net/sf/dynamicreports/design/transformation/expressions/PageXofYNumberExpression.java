package net.sf.dynamicreports.design.transformation.expressions;

import static java.util.Optional.ofNullable;

import java.text.MessageFormat;
import java.util.List;
import net.sf.dynamicreports.report.builder.expression.AbstractComplexExpression;
import net.sf.dynamicreports.report.constant.Constants;
import net.sf.dynamicreports.report.definition.ReportParameters;
import net.sf.dynamicreports.report.definition.expression.DRIExpression;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.Validate;

public class PageXofYNumberExpression extends AbstractComplexExpression<String> {
  private static final long serialVersionUID = Constants.SERIAL_VERSION_UID;

  private final int index;

  public PageXofYNumberExpression(DRIExpression<String> pageNumberFormatExpression, int index) {
    addExpression(pageNumberFormatExpression);
    this.index = index;
  }

  // CHECKSTYLE:OFF
  @Override
  public String evaluate(List<?> values, ReportParameters reportParameters) {
    String pattern = (String) values.get(0);
    Validate.isTrue(
        StringUtils.contains(pattern, "{0}"),
        "Wrong format pattern \"" + pattern + "\", missing argument {0}");
    Validate.isTrue(
        StringUtils.contains(pattern, "{1}"),
        "Wrong format pattern \"" + pattern + "\", missing argument {1}");
    Validate.isTrue(
        pattern.indexOf("{0}") < pattern.indexOf("{1}"),
        "Wrong format pattern \"" + pattern + "\", argument {0} must be before {1}");
    int index1 = pattern.indexOf("{0}");
    if (index == 0) {
      pattern = pattern.substring(0, index1 + 3);
    } else {
      pattern = pattern.substring(index1 + 3);
    }
    MessageFormat format = new MessageFormat(pattern, reportParameters.getLocale());
    Integer pageNumber = reportParameters.getPageNumber();
    return format.format(
        new Object[] {
          pageNumber,
          pageNumber,
          ofNullable(pageNumber)
              .map(n -> Math.abs(n) % 100)
              .map(n -> n % 10 == 1 && (n < 10 || n > 20) ? 1 : 2)
              .orElse(1)
        });
  }
  // CHECKSTYLE:ON
}
