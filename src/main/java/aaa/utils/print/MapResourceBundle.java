package aaa.utils.print;

import static aaa.utils.print.DynamicTemplates.FIELD_RESOURCE_PREFIX;
import static org.apache.commons.lang3.StringUtils.replaceChars;

import com.google.common.collect.Iterators;
import java.util.Enumeration;
import java.util.HashMap;
import java.util.Map;
import java.util.MissingResourceException;
import java.util.Objects;
import java.util.ResourceBundle;

public class MapResourceBundle extends ResourceBundle {

  final Map<String, String> messages = new HashMap<>();

  @Override
  protected Object handleGetObject(String key) {
    return messages.get(key);
  }

  @Override
  public Enumeration<String> getKeys() {
    return Iterators.asEnumeration(messages.keySet().iterator());
  }

  public static String deHyphenate(String template) {
    return replaceChars(template, "­", "");
  }

  public MapResourceBundle add(Map<String, String> map) {
    map.forEach((key, value) -> add(key, deHyphenate(value)));
    return this;
  }

  public MapResourceBundle fieldTitle(String fieldName, String template) {
    return add(FIELD_RESOURCE_PREFIX + fieldName, template);
  }

  public MapResourceBundle add(String key, String template) {
    return add(key, template, false);
  }

  public MapResourceBundle add(String key, String template, boolean keepHyphens) {
    messages.put(key, keepHyphens ? template.replace("­", "­\n") : deHyphenate(template));
    return this;
  }

  public String getStringSafe(String name) {
    try {
      return getString(name);
    } catch (MissingResourceException e) {
      return "";
    }
  }

  @Override
  public int hashCode() {
    return messages.hashCode();
  }

  @Override
  public boolean equals(Object obj) {
    return obj instanceof MapResourceBundle bundle && Objects.equals(messages, bundle.messages);
  }
}
