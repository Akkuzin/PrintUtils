package aaa.utils.print;

import static com.google.common.base.Strings.nullToEmpty;

import aaa.threadlocal.NotSettableThreadLocal;
import java.math.RoundingMode;
import java.text.DecimalFormat;
import java.text.DecimalFormatSymbols;
import java.text.NumberFormat;
import java.util.Locale;
import net.sf.jasperreports.engine.util.DefaultFormatFactory;

@SuppressWarnings("checkstyle:MagicNumber")
public class FormatFactory extends DefaultFormatFactory {

  private static final ThreadLocal<NumberFormat> PLAIN_NUMBER_NO_GROUP_FORMAT_HOLDER =
      NotSettableThreadLocal.withInitial(() -> makeNumberFormat(false, 0));

  private static final ThreadLocal<NumberFormat> PLAIN_NUMBER_GROUP_FORMAT_HOLDER =
      NotSettableThreadLocal.withInitial(() -> makeNumberFormat(true, 0));

  public static NumberFormat makeNumberFormat(boolean grouping, int digitsAfterComma) {
    NumberFormat result = makeDefaultNumberFormat(digitsAfterComma);
    result.setGroupingUsed(grouping);
    return result;
  }

  public static NumberFormat makeDefaultNumberFormat() {
    return makeDefaultNumberFormat(2);
  }

  public static NumberFormat makeDefaultNumberFormat(int digitsAfterComma) {
    NumberFormat result = new DecimalFormat("###,###.##", makeDecimalFormatSymbols());
    result.setMinimumFractionDigits(digitsAfterComma);
    result.setMaximumFractionDigits(digitsAfterComma);
    result.setGroupingUsed(true);
    result.setRoundingMode(RoundingMode.HALF_UP);
    return result;
  }

  public static DecimalFormatSymbols makeDecimalFormatSymbols() {
    DecimalFormatSymbols decimalFormatSymbols = new DecimalFormatSymbols();
    decimalFormatSymbols.setDecimalSeparator('.');
    decimalFormatSymbols.setGroupingSeparator(' ');
    return decimalFormatSymbols;
  }

  public NumberFormat createNumberFormat(String pattern, Locale locale) {
    return switch (nullToEmpty(pattern)) {
      case PLAIN_NUMBER_PATTERN -> PLAIN_NUMBER_GROUP_FORMAT_HOLDER.get();
      case PLAIN_NUMBER_NO_GROUP_PATTERN -> PLAIN_NUMBER_NO_GROUP_FORMAT_HOLDER.get();
      default -> super.createNumberFormat(pattern, locale);
    };
  }

  public static final String PLAIN_NUMBER_PATTERN = "#,##0";
  public static final String PLAIN_NUMBER_NO_GROUP_PATTERN = "###0";
}
