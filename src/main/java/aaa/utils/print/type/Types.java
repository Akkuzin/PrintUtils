package aaa.utils.print.type;

public class Types {

  private static final LocalDateType localDateType = new LocalDateType();
  private static final LocalDateTimeType localDateTimeType = new LocalDateTimeType();

  public static LocalDateType localDateType() {
    return localDateType;
  }

  public static LocalDateTimeType localDateTimeType() {
    return localDateTimeType;
  }
}
