package aaa.utils.poi.formula;

import lombok.AccessLevel;
import lombok.AllArgsConstructor;
import lombok.experimental.FieldDefaults;

@FieldDefaults(level = AccessLevel.PRIVATE, makeFinal = true)
@AllArgsConstructor
public class MultiplyFormula extends BaseFormula {

	protected static final String MULTIPLY_FORMULA_NAME = "*";

	IBaseFormula firstMultiplier;
	IBaseFormula secondMultiplier;

	@Override
	public String formatAsString() {
		return surroundWithBrackets(firstMultiplier.formatAsString()) + MULTIPLY_FORMULA_NAME
			+ surroundWithBrackets(secondMultiplier.formatAsString());
	}

}
