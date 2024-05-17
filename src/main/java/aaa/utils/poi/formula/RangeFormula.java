package aaa.utils.poi.formula;

import lombok.AccessLevel;
import lombok.experimental.FieldDefaults;
import aaa.utils.poi.FormulaPoiUtils;

@FieldDefaults(level = AccessLevel.PRIVATE, makeFinal = true)
public class RangeFormula extends BaseFormula {

	protected static final String RANGE_FORMULA_DELIMITER = ":";

	String range;

	public RangeFormula(String cellStart) {
		this.range = cellStart;
	}

	public RangeFormula(String cellStart, String cellFinish) {
		this.range = cellStart + RANGE_FORMULA_DELIMITER + cellFinish;
	}

	public RangeFormula(long rowStart, long columnStart) {
		this(FormulaPoiUtils.getCellAddress(rowStart, columnStart));
	}

	public RangeFormula(long rowStart, long columnStart, long rowFinish, long columnFinish) {
		this(	FormulaPoiUtils.getCellAddress(rowStart, columnStart),
				FormulaPoiUtils.getCellAddress(rowFinish, columnFinish));
	}

	@Override
	public String formatAsString() {
		return range;
	}

}
