import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.ToString;

@ToString
@Getter
@NoArgsConstructor
@AllArgsConstructor
public class Book implements ExcelRow {

    @ExcelColumn(cell = "A")
    private String No;

    @ExcelColumn(cell = "B")
    private String title;

    @ExcelColumn(cell = "C")
    private Double price;

}
