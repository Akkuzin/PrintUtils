package aaa.utils.poi;

import static java.util.Arrays.asList;
import static org.junit.jupiter.api.Assertions.*;

import aaa.utils.poi.ExcelExport.CellDefinition;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.time.LocalDateTime;
import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.SneakyThrows;
import org.junit.jupiter.api.Test;

class ExcelExportTest {

  @Getter
  @AllArgsConstructor
  static class Data {
    Long id;
    String name;
    LocalDateTime moment;
  }

  @SneakyThrows
  @Test
  public void test() {
    Files.write(
        Paths.get("./target/excelExport.xlsx"),
        ExcelExport.makeReport(
            asList(new Data(1L, "Info", LocalDateTime.now())),
            asList(
                CellDefinition.forGetter(Data::getId).title("ID").weight(10).build(),
                CellDefinition.forGetter(Data::getName).title("Name").weight(30).build(),
                CellDefinition.forGetter(Data::getMoment)
                    .title("Дата без времени")
                    .dateWithTime(false)
                    .weight(20)
                    .build(),
                CellDefinition.forGetter(Data::getMoment)
                    .dateWithTime(true)
                    .title("Дата со временем")
                    .weight(20)
                    .build())));
  }
}
