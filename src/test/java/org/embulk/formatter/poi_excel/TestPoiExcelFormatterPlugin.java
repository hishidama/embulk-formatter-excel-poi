package org.embulk.formatter.poi_excel;

import static org.junit.Assert.assertEquals;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.UncheckedIOException;
import java.time.Instant;
import java.time.LocalDate;
import java.time.LocalTime;
import java.time.ZoneId;
import java.time.ZonedDateTime;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.embulk.config.ConfigSource;
import org.junit.Test;

import com.hishidama.embulk.tester.EmbulkPluginTester;
import com.hishidama.embulk.tester.EmbulkTestParserConfig;

public class TestPoiExcelFormatterPlugin {

    @Test
    public void test97_sheetNameNull() {
        test("EXCEL97", null);
    }

    @Test
    public void test97() {
        test("EXCEL97", "s");
    }

    @Test
    public void test2007_sheetNameNull() {
        test("EXCEL2007", null);
    }

    @Test
    public void test2007() {
        test("EXCEL2007", "s");
    }

    private void test(String spreadSheetVersion, String sheetName) {
        try (EmbulkPluginTester tester = new EmbulkPluginTester()) {
            tester.addFormatterPlugin(PoiExcelFormatterPlugin.TYPE, PoiExcelFormatterPlugin.class);
            final int OUTPUT_TASK_SIZE = 1;
            tester.setOutputMinTaskSize(OUTPUT_TASK_SIZE);

            List<String> inList = Arrays.asList( //
                    "false,11,12,abc,2023/05/13", //
                    "true,21,22,def,2023/05/14");
            EmbulkTestParserConfig parser = tester.newParserConfig("csv");
            parser.set("stop_on_invalid_record", true);
            parser.set("default_timezone", "Asia/Tokyo");
            parser.set("default_timestamp_format", "%Y/%m/%d");
            parser.addColumn("bool", "boolean");
            parser.addColumn("num1", "long");
            parser.addColumn("num2", "double");
            parser.addColumn("text", "string");
            parser.addColumn("date", "timestamp");

            ConfigSource columnOptions = tester.newConfigSource();
            columnOptions.set("date", "{data_format: \"yyyy/mm/dd\"}");

            ConfigSource formatter = tester.newConfigSource();
            formatter.set("type", "poi_excel");
            formatter.set("spread_sheet_version", spreadSheetVersion);
            if (sheetName != null) {
                formatter.set("sheet_name", sheetName);
            }
            formatter.set("column_options", columnOptions);

            List<byte[]> resultList = tester.runFormatterToBinary(inList, parser, formatter);

            assertEquals(OUTPUT_TASK_SIZE, resultList.size());
            byte[] result = resultList.get(0);

            Workbook workbook = createWorkbook(result);
            Sheet sheet = workbook.getSheet((sheetName != null) ? sheetName : "Sheet1");
            assertEquals(0, sheet.getFirstRowNum());
            assertEquals(inList.size() - 1, sheet.getLastRowNum());
            {
                Row row = sheet.getRow(0);
                int i = 0;
                assertEquals(false, row.getCell(i++).getBooleanCellValue());
                assertEquals(11, row.getCell(i++).getNumericCellValue(), 0);
                assertEquals(12, row.getCell(i++).getNumericCellValue(), 0);
                assertEquals("abc", row.getCell(i++).getStringCellValue());
                assertEqualsDate(2023, 5, 13, row.getCell(i++).getDateCellValue());
            }
            {
                Row row = sheet.getRow(1);
                int i = 0;
                assertEquals(true, row.getCell(i++).getBooleanCellValue());
                assertEquals(21, row.getCell(i++).getNumericCellValue(), 0);
                assertEquals(22, row.getCell(i++).getNumericCellValue(), 0);
                assertEquals("def", row.getCell(i++).getStringCellValue());
                assertEqualsDate(2023, 5, 14, row.getCell(i++).getDateCellValue());
            }
        }
    }

    private static Workbook createWorkbook(byte[] in) {
        try (InputStream is = new ByteArrayInputStream(in)) {
            return WorkbookFactory.create(is);
        } catch (IOException e) {
            throw new UncheckedIOException(e.getMessage(), e);
        }
    }

    private static void assertEqualsDate(int year, int month, int day, Date actualDate) {
        ZoneId zoneId = ZoneId.of("Asia/Tokyo");
        Instant expected = ZonedDateTime.of(LocalDate.of(year, month, day), LocalTime.MIN, zoneId).toInstant();
        assertEquals(expected, actualDate.toInstant());
    }
}
