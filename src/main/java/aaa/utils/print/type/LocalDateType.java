package aaa.utils.print.type;

import static aaa.utils.print.DynamicUtils.valueFormatter;

import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.time.temporal.Temporal;
import java.util.Locale;
import net.sf.dynamicreports.report.base.datatype.AbstractDataType;
import net.sf.dynamicreports.report.constant.Constants;
import net.sf.dynamicreports.report.constant.HorizontalTextAlignment;
import net.sf.dynamicreports.report.defaults.Defaults;
import net.sf.dynamicreports.report.definition.expression.DRIValueFormatter;
import net.sf.dynamicreports.report.exception.DRException;

public class LocalDateType extends AbstractDataType<Temporal, Temporal> {
  private static final long serialVersionUID = Constants.SERIAL_VERSION_UID;

  @Override
  public String getPattern() {
    return Defaults.getDefaults().getDateType().getPattern();
  }

  @Override
  public HorizontalTextAlignment getHorizontalTextAlignment() {
    return Defaults.getDefaults().getDateType().getHorizontalTextAlignment();
  }

  @Override
  public DRIValueFormatter<String, Temporal> getValueFormatter() {
    return valueFormatter(DateTimeFormatter.ofPattern(getPattern())::format);
  }

  @Override
  public String valueToString(Temporal value, Locale locale) {
    if (value != null) {
      return getValueFormatter().format(value, null);
    }
    return null;
  }

  @Override
  public LocalDate stringToValue(String value, Locale locale) throws DRException {
    if (value != null) {
      try {
        return LocalDate.from(DateTimeFormatter.ofPattern(getPattern(), locale).parse(value));
      } catch (DateTimeParseException e) {
        throw new DRException("Unable to convert string value to date", e);
      }
    }
    return null;
  }
}
