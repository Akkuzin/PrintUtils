package aaa.utils.print;

import static aaa.utils.print.type.Types.localDateTimeType;
import static aaa.utils.print.type.Types.localDateType;
import static net.sf.dynamicreports.report.builder.DynamicReports.col;

import jakarta.persistence.metamodel.Attribute;
import java.time.temporal.Temporal;
import lombok.experimental.UtilityClass;
import net.sf.dynamicreports.report.builder.DynamicReports;
import net.sf.dynamicreports.report.builder.FieldBuilder;
import net.sf.dynamicreports.report.builder.column.TextColumnBuilder;
import net.sf.dynamicreports.report.definition.datatype.DRIDataType;
import net.sf.dynamicreports.report.definition.expression.DRIExpression;

@UtilityClass
public class DynamicJpaUtils {

  public static <T> TextColumnBuilder<T> column(String title, Attribute<?, T> field) {
    return col.column(title, field.getName(), field.getJavaType());
  }

  public static <X, T> TextColumnBuilder<T> column(
      String title, Attribute<?, X> field, Attribute<X, T> subfield) {
    return col.column(title, field.getName() + "." + subfield.getName(), subfield.getJavaType());
  }

  public static <T extends Temporal> TextColumnBuilder<T> columnDate(
      String title, Attribute<?, T> field) {
    return col.column(title, field.getName(), (DRIDataType<? super T, T>) localDateType());
  }

  public static <T extends Temporal> TextColumnBuilder<T> columnDateTime(
      String title, Attribute<?, T> field) {
    return col.column(title, field.getName(), (DRIDataType<? super T, T>) localDateTimeType());
  }

  public static DRIExpression<String> columnTitle(Attribute<?, ?> attribute) {
    return DynamicTemplates.columnTitle(attribute.getName());
  }

  public static <T> FieldBuilder<T> field(Attribute<?, T> field) {
    return DynamicReports.field(field.getName(), field.getJavaType());
  }

  public static <T extends Temporal> FieldBuilder<T> fieldDate(Attribute<?, T> field) {
    return DynamicReports.field(field.getName(), (DRIDataType<? super T, T>) localDateType());
  }

  public static <T extends Temporal> FieldBuilder<T> fieldDateTime(Attribute<?, T> field) {
    return DynamicReports.field(field.getName(), (DRIDataType<? super T, T>) localDateTimeType());
  }

  public static <X, T> FieldBuilder<T> field(Attribute<?, X> field, Attribute<X, T> subfield) {
    return DynamicReports.field(field.getName() + "." + subfield.getName(), subfield.getJavaType());
  }
}
