package aaa.utils.poi.formula;

import aaa.nvl.Nvl;
import lombok.AccessLevel;
import lombok.NoArgsConstructor;
import lombok.experimental.FieldDefaults;
import org.apache.commons.lang3.StringUtils;

import java.util.ArrayList;
import java.util.List;

import static java.util.stream.Collectors.joining;
import static java.util.stream.Collectors.toList;

@NoArgsConstructor
@FieldDefaults(level = AccessLevel.PRIVATE, makeFinal = true)
public class SumFormula extends BaseFormula {

	protected static final String SUM_FORMULA_NAME = "SUM";
	protected static final String SUM_FORMULA_DELIMITER = ",";
	protected static final int SUM_MAX_OPERANDS = 30;

	List<IBaseFormula> list = new ArrayList<>();

	public SumFormula(IBaseFormula formula) {
		this();
		add(formula);
	}

	private List<String> prepareOperands(List<String> operands) {
		List<String> result;
		int operandsSize = operands.size();
		if (operandsSize <= SUM_MAX_OPERANDS) {
			result = operands;
		} else {
			result = new ArrayList<>((operandsSize - 1) % SUM_MAX_OPERANDS + 1);
			for (int i = 0; i < operandsSize; i += SUM_MAX_OPERANDS) {
				result.add(formatAsString(operands.subList(i, Nvl.min(	i + SUM_MAX_OPERANDS,
																		operandsSize))));
			}
			result = prepareOperands(result);
		}
		return result;
	}

	private String formatAsString(List<String> operands) {
		return SUM_FORMULA_NAME
			+ surroundWithBrackets(prepareOperands(operands).stream()
					.filter(StringUtils::isNotBlank)
					.collect(joining(SUM_FORMULA_DELIMITER)));
	}

	@Override
	public String formatAsString() {
		return formatAsString(list.stream().map(IBaseFormula::formatAsString).collect(toList()));
	}

	public void add(IBaseFormula expression) {
		list.add(expression);
	}

	public void add(List<IBaseFormula> expressions) {
		list.addAll(expressions);
	}

}
