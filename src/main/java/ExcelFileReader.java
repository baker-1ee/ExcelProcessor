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
import java.util.function.Consumer;
import java.util.function.Supplier;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ExcelFileReader<T extends ExcelRow> implements XSSFSheetXMLHandler.SheetContentsHandler {

    private final Supplier<T> instanceGenerator;
    private final Consumer<List<T>> processor;

    private final LinkedHashMap<String, Field> fieldMap = new LinkedHashMap<>();

    private int currentRowNum = 0;
    private T currentRow;
    private List<T> partitionRows = new ArrayList<>();
    private final List<String> header = new ArrayList<>();

    public ExcelFileReader(Supplier<T> instanceGenerator, Consumer<List<T>> processor) {
        this.instanceGenerator = instanceGenerator;
        this.processor = processor;
        setFieldMap();
    }

    private void setFieldMap() {
        ExcelRow excelRow = this.instanceGenerator.get();
        Field[] fields = excelRow.getClass().getDeclaredFields();
        for (Field field : fields) {
            ExcelColumn annotation = field.getAnnotation(ExcelColumn.class);
            if (annotation == null) continue;
            String cell = annotation.cell();
            this.fieldMap.put(cell, field);
        }
    }

    public void processExcel(File file) throws Exception {
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
        if(isHeader()) return;
        this.partitionRows.add(this.currentRow);
        int PARTITION_SIZE = 2;
        if(this.partitionRows.size() == PARTITION_SIZE) {
            this.processor.accept(this.partitionRows);
            this.partitionRows = new ArrayList<>();
        }
    }

    @SneakyThrows
    @Override
    public void cell(String columnName, String value, XSSFComment comment) {
        if (isHeader()) {
            this.header.add(value);
            return;
        }

        StringBuilder cellAlphabet = getCellAlphabet(columnName);
        Field field = this.currentRow.getClass().getDeclaredField(this.fieldMap.get(cellAlphabet.toString()).getName());
        field.setAccessible(true);

        if (field.getType().equals(String.class)) field.set(this.currentRow, value);
        else if (field.getType().equals(Double.class)) field.set(this.currentRow, Double.valueOf(value));
        else throw new RuntimeException("field type not found");
    }

    private StringBuilder getCellAlphabet(String columnName) {
        Pattern pattern = Pattern.compile("[a-zA-Z]");
        Matcher matcher = pattern.matcher(columnName);
        StringBuilder cellAlphabet = new StringBuilder();
        while (matcher.find()) {
            cellAlphabet.append(matcher.group());
        }
        return cellAlphabet;
    }

    private boolean isHeader() {
        return this.currentRowNum == 0;
    }

    @Override
    public void endSheet() {
        this.processor.accept(this.partitionRows);
    }

    public List<String> getHeader() {
        return this.header;
    }
}
