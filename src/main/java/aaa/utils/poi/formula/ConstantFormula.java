package aaa.utils.poi.formula;

import lombok.AccessLevel;
import lombok.AllArgsConstructor;
import lombok.experimental.FieldDefaults;

import static aaa.nvl.Nvl.nvlToString;

@FieldDefaults(level = AccessLevel.PRIVATE, makeFinal = true)
@AllArgsConstructor
public class ConstantFormula extends BaseFormula {

	String value;

	public <T> ConstantFormula(T value) {
		this(nvlToString(value, ""));
	}

	@Override
	public String formatAsString() {
		return value;
	}

}
