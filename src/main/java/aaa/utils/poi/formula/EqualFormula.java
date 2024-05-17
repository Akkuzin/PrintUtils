package aaa.utils.poi.formula;

import lombok.AccessLevel;
import lombok.AllArgsConstructor;
import lombok.experimental.FieldDefaults;

@FieldDefaults(level = AccessLevel.PRIVATE, makeFinal = true)
@AllArgsConstructor
public class EqualFormula extends BaseFormula {

	protected static final String EQUAL_FORMULA_NAME = "=";

	IBaseFormula firstFormula;
	IBaseFormula secondFormula;

	@Override
	public String formatAsString() {
		return surroundWithBrackets(firstFormula.formatAsString()) + EQUAL_FORMULA_NAME
			+ surroundWithBrackets(secondFormula.formatAsString());
	}

}
