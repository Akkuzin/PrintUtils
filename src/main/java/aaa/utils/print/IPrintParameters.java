package aaa.utils.print;

import static aaa.nvl.Nvl.nvl;

import aaa.i18n.I18NResolver;

public interface IPrintParameters {

  Format getFormat();

  default Format getFormatSafe() {
    return nvl(getFormat(), Format.PDF);
  }

  default I18NResolver getI18n() {
    return I18NResolver.DEFAULT_I18N;
  }

  default I18NResolver getI18nSafe() {
    return nvl(getI18n(), I18NResolver.DEFAULT_I18N);
  }
}
