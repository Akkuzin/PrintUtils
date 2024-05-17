package aaa.utils.print;

import com.google.common.collect.ImmutableList;
import java.util.List;
import lombok.AccessLevel;
import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.experimental.FieldDefaults;

@AllArgsConstructor(access = AccessLevel.PRIVATE)
@Getter
@FieldDefaults(level = AccessLevel.PUBLIC, makeFinal = true)
public enum Format {
  PDF(true, "pdf", "application/pdf"),
  XLS(false, "xls", "application/vnd.ms-excel"),
  XLS_PAGES(true, "xls", "application/vnd.ms-excel"),
  XLSX(false, "xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
  XLSX_PAGES(true, "xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
  HTML(false, "html", "text/html"),
  RTF(true, "rtf", "application/rtf"),
  DOC(true, "doc", "application/msword"),
  DOCX(true, "docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"),
  TXT(false, "txt", "text/plain"),
  XML(false, "xml", "application/xml"),
  MDB(false, "mdb", "application/octet-stream"),
  CSV(false, "csv", "text/csv");

  boolean defaultPaging;
  String extention;
  String mimeType;

  public static final List<Format> ALL_EXCEL = ImmutableList.of(XLS, XLS_PAGES, XLSX, XLSX_PAGES);
  public static final List<Format> ALL_WORD = ImmutableList.of(DOC, DOCX);

  public static boolean isExcel(Format format) {
    return ALL_EXCEL.contains(format);
  }

  public static boolean isWord(Format format) {
    return ALL_WORD.contains(format);
  }

  @SuppressWarnings("checkstyle:AbbreviationAsWordInName")
  public static boolean isOOXML(Format format) {
    return XLSX == format || XLSX_PAGES == format;
  }
}
