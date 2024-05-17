package aaa.utils.poi.formula;

import static java.util.stream.Collectors.joining;

import java.util.stream.Stream;
import lombok.AccessLevel;
import lombok.AllArgsConstructor;
import lombok.experimental.FieldDefaults;

@FieldDefaults(level = AccessLevel.PRIVATE, makeFinal = true)
@AllArgsConstructor
public class IfFormula extends BaseFormula {

  protected static final String IF_FORMULA_NAME = "IF";
  protected static final String IF_FORMULA_DELIMITER = ",";

  IBaseFormula condition;
  IBaseFormula valueIfTrue;
  IBaseFormula valueIfFalse;

  @Override
  public String formatAsString() {
    return IF_FORMULA_NAME
        + surroundWithBrackets(
            Stream.of(condition, valueIfTrue, valueIfFalse)
                .map(IBaseFormula::formatAsString)
                .collect(joining(IF_FORMULA_DELIMITER)));
  }
}
