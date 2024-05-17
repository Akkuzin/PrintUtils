package aaa.utils.print;

import java.util.function.BiFunction;
import java.util.function.Function;
import lombok.experimental.UtilityClass;
import net.sf.dynamicreports.report.base.expression.AbstractSimpleExpression;
import net.sf.dynamicreports.report.base.expression.AbstractValueFormatter;
import net.sf.dynamicreports.report.definition.ReportParameters;

@UtilityClass
public class DynamicUtils {

  public static <T> AbstractSimpleExpression<T> simpleExpression(
      Class<T> clazz, Function<ReportParameters, T> supplier) {
    return new AbstractSimpleExpression<>() {

      @Override
      public T evaluate(ReportParameters reportParameters) {
        return supplier.apply(reportParameters);
      }

      @Override
      public Class<? super T> getValueClass() {
        return clazz;
      }
    };
  }

  public static <T, V> AbstractValueFormatter<V, T> valueFormatter(
      Class<V> clazz, Function<T, V> function) {
    return new AbstractValueFormatter<>() {

      @Override
      public V format(T value, ReportParameters reportParameters) {
        return value == null ? null : function.apply(value);
      }

      @Override
      public Class<V> getValueClass() {
        return clazz;
      }
    };
  }

  public static <T> AbstractValueFormatter<String, T> valueFormatter(Function<T, String> function) {
    return valueFormatter(String.class, function);
  }

  public static <T, V> AbstractValueFormatter<V, T> valueFormatter(
      Class<V> clazz, BiFunction<T, ReportParameters, V> function) {
    return new AbstractValueFormatter<>() {

      @Override
      public V format(T value, ReportParameters reportParameters) {
        return value == null ? null : function.apply(value, reportParameters);
      }

      @Override
      public Class<V> getValueClass() {
        return clazz;
      }
    };
  }

  public static <T> AbstractValueFormatter<String, T> valueFormatter(
      BiFunction<T, ReportParameters, String> function) {
    return valueFormatter(String.class, function);
  }
}
