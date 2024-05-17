package aaa.utils.print;

public interface IPrint<P extends IPrintParameters> {

  byte[] generateBlank(P params);
}
