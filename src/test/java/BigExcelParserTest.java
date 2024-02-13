import lombok.*;
import lombok.extern.slf4j.Slf4j;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.springframework.test.context.junit.jupiter.SpringExtension;

import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.function.Consumer;

import static org.assertj.core.api.Assertions.assertThat;

@Slf4j
class BigExcelParserTest {

    @Test
    @DisplayName("ExcelFileReader 기본 사용법")
    @SneakyThrows
    void excelReadTest() {
        // given
        String filePath = "src/main/resources/simple.xlsx";
        File file = new File(filePath);
        List<Book> parsedBookList = new ArrayList<>();

        // when
        BigExcelParser<Book> bigExcelParser = new BigExcelParser<>(Book::new);
        bigExcelParser.processExcel(file, myProcessor(parsedBookList));

        // then
        assertThat(parsedBookList.size()).isEqualTo(5);
    }

    // Excel Data 를 Chunk Size 단위로 파싱한 뒤 수행 시키고 싶은 로직을 Consumer 로 작성
    private Consumer<List<Book>> myProcessor(List<Book> parsedBookList) {
        return (books) -> {
            log.info("partition process : {}", books.toString());
            parsedBookList.addAll(books);
        };
    }

    @Test
    @DisplayName("Excel Data 파싱 후에는 컬럼의 header 목록도 별도로 얻을 수 있다.")
    @SneakyThrows
    void getHeaderTest() {
        // given
        String filePath = "src/main/resources/simple.xlsx";
        File file = new File(filePath);

        BigExcelParser<Book> bigExcelParser = new BigExcelParser<>(Book::new);
        bigExcelParser.processExcel(file, (books) -> {});

        // when
        List<String> header = bigExcelParser.getHeader();

        // then
        assertThat(header.size()).isEqualTo(3);
        assertThat(header.get(0)).isEqualTo("번호");
        assertThat(header.get(1)).isEqualTo("도서명");
        assertThat(header.get(2)).isEqualTo("가격");
    }

    @Data
    @NoArgsConstructor
    private static class Book implements ExcelRow {

        @ExcelColumn(cellAlphabet = "A")
        private String No;

        @ExcelColumn(cellAlphabet = "B")
        private String title;

        @ExcelColumn(cellAlphabet = "C")
        private Double price;

    }

}
