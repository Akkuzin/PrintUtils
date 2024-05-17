package aaa.utils.print;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.experimental.FieldDefaults;
import lombok.experimental.SuperBuilder;

@Data
@SuperBuilder
@AllArgsConstructor(staticName = "of")
@FieldDefaults(makeFinal = true)
public class ReportType {

  Format format;
}
