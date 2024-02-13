import lombok.SneakyThrows;
import org.apache.poi.ooxml.util.SAXHelper;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;

import java.io.File;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.function.Consumer;
import java.util.function.Supplier;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class BigExcelParser<T extends ExcelRow> implements XSSFSheetXMLHandler.SheetContentsHandler {

    private static final int PARTITION_SIZE = 2;
    private static final int DATA_START_ROW_NUM = 1;

    private final Supplier<T> instanceGenerator;
    private final Map<String, Field> fieldMap = new LinkedHashMap<>();
    private final List<String> header = new ArrayList<>();
    private Consumer<List<T>> processor;
    private int currentRowNum = 0;

    private T currentRow;
    private List<T> partitionRows = new ArrayList<>();

    /**
     * Excel File Reader 생성
     *
     * @param instanceGenerator Excel Data 파싱 후 담을 객체의 instance 생성기
     */
    public BigExcelParser(Supplier<T> instanceGenerator) {
        this.instanceGenerator = instanceGenerator;
        setFieldMap();
    }

    /**
     * Excel 의 Cell 과 ExcelRow class 의 Field 맵 생성
     */
    private void setFieldMap() {
        ExcelRow excelRow = this.instanceGenerator.get();
        Field[] fields = excelRow.getClass().getDeclaredFields();
        for (Field field : fields) {
            ExcelColumn annotation = field.getAnnotation(ExcelColumn.class);
            if (annotation == null) continue;
            String cell = annotation.cellAlphabet();
            this.fieldMap.put(cell, field);
        }
    }

    /**
     * Excel File 을 Streaming방식 으로 읽으면서 Chunk 단위로 사용자 정의 처리기를 수행
     *
     * @param file      파싱 대상 Excel File
     * @param processor Excel Data 파싱 후 Chunk 단위로 수행할 사용자 정의 처리기
     * @throws Exception 예외
     */
    public void processExcel(File file, Consumer<List<T>> processor) throws Exception {
        this.processor = processor;

        // SAX (Simple API for XML) Library 사용을 위한 boilerplate code
        OPCPackage opcPackage = OPCPackage.open(file);
        XSSFReader xssfReader = new XSSFReader(opcPackage);
        StylesTable stylesTable = xssfReader.getStylesTable();
        ReadOnlySharedStringsTable stringsTable = new ReadOnlySharedStringsTable(opcPackage);

        InputStream inputStream = xssfReader.getSheetsData().next();
        InputSource inputSource = new InputSource(inputStream);
        ContentHandler handler = new XSSFSheetXMLHandler(stylesTable, stringsTable, this, false);

        XMLReader xmlReader = SAXHelper.newXMLReader();
        xmlReader.setContentHandler(handler);

        xmlReader.parse(inputSource);
        inputStream.close();
        opcPackage.close();
    }

    @Override
    public void startRow(int rowNum) {
        this.currentRowNum = rowNum;
        this.currentRow = this.instanceGenerator.get();
    }

    @Override
    public void endRow(int rowNum) {
        if (isHeader()) return;
        this.partitionRows.add(this.currentRow);
        if (this.partitionRows.size() == PARTITION_SIZE) {
            this.processor.accept(this.partitionRows);
            this.partitionRows = new ArrayList<>();
        }
    }

    @SneakyThrows
    @Override
    // cell : Excel 의 Cell 이름 (e.g. A1, A2)
    // value : Cell 의 값
    public void cell(String cell, String value, XSSFComment comment) {
        if (isHeader()) {
            this.header.add(value);
            return;
        }

        StringBuilder cellAlphabet = getCellAlphabet(cell);
        Field field = this.currentRow.getClass().getDeclaredField(this.fieldMap.get(cellAlphabet.toString()).getName());
        field.setAccessible(true);

        if (field.getType().equals(String.class)) field.set(this.currentRow, value);
        else if (field.getType().equals(Double.class)) field.set(this.currentRow, Double.valueOf(value));
        else throw new RuntimeException("field type not found");
    }

    // Cell (e.g. A1, A2) 에서 알파벳만 추출 (e.g. A, A)
    private StringBuilder getCellAlphabet(String columnName) {
        Pattern pattern = Pattern.compile("[a-zA-Z]");
        Matcher matcher = pattern.matcher(columnName);
        StringBuilder cellAlphabet = new StringBuilder();
        while (matcher.find()) {
            cellAlphabet.append(matcher.group());
        }
        return cellAlphabet;
    }

    // Excel Data 의 Header 여부
    private boolean isHeader() {
        return this.currentRowNum == DATA_START_ROW_NUM - 1;
    }

    @Override
    public void endSheet() {
        this.processor.accept(this.partitionRows);
    }

    /**
     * Excel Data 파싱 후 header 목록이 채워짐
     */
    public List<String> getHeader() {
        return this.header;
    }
}
