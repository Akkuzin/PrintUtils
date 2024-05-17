package aaa.utils.print.expression;

import lombok.AllArgsConstructor;
import lombok.experimental.FieldDefaults;
import net.sf.dynamicreports.report.base.expression.AbstractSimpleExpression;
import net.sf.dynamicreports.report.definition.ReportParameters;
import net.sf.jasperreports.engine.JRDataSource;
import net.sf.jasperreports.engine.data.JRBeanCollectionDataSource;

@AllArgsConstructor(staticName = "of")
@FieldDefaults(makeFinal = true)
public class SubreportFieldDatasourceExpression extends AbstractSimpleExpression<JRDataSource> {

  String fieldName;

  @Override
  public JRDataSource evaluate(ReportParameters reportParameters) {
    return new JRBeanCollectionDataSource(reportParameters.getFieldValue(fieldName));
  }
}
