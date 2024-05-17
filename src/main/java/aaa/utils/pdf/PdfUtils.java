package aaa.utils.pdf;

import com.lowagie.text.Image;
import com.lowagie.text.Rectangle;
import com.lowagie.text.pdf.PdfReader;
import com.lowagie.text.pdf.PdfStamper;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Collection;
import lombok.SneakyThrows;
import lombok.experimental.UtilityClass;
import org.apache.pdfbox.Loader;
import org.apache.pdfbox.io.MemoryUsageSetting;
import org.apache.pdfbox.io.RandomAccessRead;
import org.apache.pdfbox.io.RandomAccessReadBuffer;
import org.apache.pdfbox.multipdf.PDFMergerUtility;
import org.apache.pdfbox.multipdf.PDFMergerUtility.DocumentMergeMode;
import org.apache.pdfbox.pdmodel.PDDocument;

@UtilityClass
public class PdfUtils {

  public static final int MERGE_MAX_MAIN_MEMORY_BYTES = 10 * 1024 * 1024;

  public static byte[] doMerge(Collection<? extends RandomAccessRead> list) {
    ByteArrayOutputStream result = new ByteArrayOutputStream();
    doMerge(result, list);
    return result.toByteArray();
  }

  @SneakyThrows
  public static byte[] doMergeByteArrays(Collection<byte[]> list) {
    try (var result = new ByteArrayOutputStream()) {
      doMergeByteArrays(result, list);
      return result.toByteArray();
    }
  }

  @SneakyThrows
  public static void doMerge(
      OutputStream outputStream, Collection<? extends RandomAccessRead> list) {
    PDFMergerUtility ut = new PDFMergerUtility();
    list.forEach(ut::addSource);
    ut.setDocumentMergeMode(DocumentMergeMode.OPTIMIZE_RESOURCES_MODE);
    ut.setDestinationStream(outputStream);
    ut.mergeDocuments(MemoryUsageSetting.setupMixed(MERGE_MAX_MAIN_MEMORY_BYTES).streamCache);
  }

  @SneakyThrows
  public static void doMergeByteArrays(OutputStream outputStream, Collection<byte[]> list) {
    doMerge(outputStream, list.stream().map(RandomAccessReadBuffer::new).toList());
  }

  public static Integer countPages(RandomAccessRead data) {
    if (data == null) {
      return null;
    }
    try (PDDocument document =
        Loader.loadPDF(
            data, MemoryUsageSetting.setupMixed(MERGE_MAX_MAIN_MEMORY_BYTES).streamCache)) {
      return document.getNumberOfPages();
    } catch (IOException e) {
      return 0;
    }
  }

  @SuppressWarnings("checkstyle:MagicNumber")
  public static byte[] stampPdfOnPlace(byte[] pdf, byte[] stamp, int place, int rowCount) {
    return stampPdf(pdf, stamp, 1f / 10, 1f * place / rowCount);
  }

  /** Установка изображения штампа на документ PDF */
  @SneakyThrows
  @SuppressWarnings("checkstyle:MagicNumber")
  public static byte[] stampPdf(byte[] pdf, byte[] stamp, Float x, Float y) {
    PdfReader reader = new PdfReader(pdf);
    try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
      PdfStamper stamper = new PdfStamper(reader, bos);
      Image img = Image.getInstance(stamp);
      int lastPage = reader.getNumberOfPages();
      Rectangle pageSize = reader.getPageSizeWithRotation(lastPage);
      img.scaleToFit(pageSize.getWidth() * 85 / 100, pageSize.getHeight() / 12);
      img.setAbsolutePosition(pageSize.getWidth() * x, pageSize.getHeight() * y);
      stamper.getOverContent(lastPage).addImage(img);
      stamper.close();
      return bos.toByteArray();
    }
  }
}
