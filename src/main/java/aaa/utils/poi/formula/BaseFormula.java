package aaa.utils.poi.formula;

public abstract class BaseFormula implements IBaseFormula {

	protected static final String OPEN_BRACKET = "(";
	protected static final String CLOSE_BRACKET = ")";

	protected static String surroundWithBrackets(String value) {
		return OPEN_BRACKET + value + CLOSE_BRACKET;
	}

}
