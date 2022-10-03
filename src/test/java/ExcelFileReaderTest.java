import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.springframework.test.context.junit.jupiter.SpringExtension;

import java.io.File;
import java.util.List;
import java.util.function.Consumer;

import static org.assertj.core.api.Assertions.assertThat;

@Slf4j
@ExtendWith(SpringExtension.class)
class ExcelFileReaderTest {

    @Test
    @SneakyThrows
    void excelReadTest() {
        String filePath = "src/main/resources/simple.xlsx";
        File file = new File(filePath);

        ExcelFileReader<Book> excelFileReader = new ExcelFileReader<>(Book::new, myProcessor());
        excelFileReader.processExcel(file);
        List<String> header = excelFileReader.getHeader();
        assertThat(header.size()).isEqualTo(3);
    }

    private Consumer<List<Book>> myProcessor() {
        return (books) -> {
            log.info("partition process : {}", books.toString());
        };
    }

}