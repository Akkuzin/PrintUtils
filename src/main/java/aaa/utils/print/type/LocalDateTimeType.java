package aaa.utils.print.type;

import net.sf.dynamicreports.report.defaults.Defaults;

public class LocalDateTimeType extends LocalDateType {

  @Override
  public String getPattern() {
    return Defaults.getDefaults().getDateType().getPattern()
        + " "
        + Defaults.getDefaults().getTimeHourToMinuteType().getPattern();
  }
}
