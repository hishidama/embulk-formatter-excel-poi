package org.embulk.formatter.poi_excel.visitor;

import java.time.Instant;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.Optional;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellUtil;
import org.embulk.formatter.poi_excel.PoiExcelFormatterPlugin.ColumnOption;
import org.embulk.formatter.poi_excel.PoiExcelFormatterPlugin.PluginTask;
import org.embulk.spi.Column;
import org.embulk.spi.ColumnVisitor;
import org.embulk.spi.PageReader;
import org.embulk.spi.json.JsonValue;

public class PoiExcelColumnVisitor implements ColumnVisitor {

    private final PluginTask task;
    private final Sheet sheet;
    private final PageReader pageReader;

    private int rowIndex = 0;

    private Row currentRow = null;

    public PoiExcelColumnVisitor(PluginTask task, Sheet sheet, PageReader pageReader) {
        this.task = task;
        this.sheet = sheet;
        this.pageReader = pageReader;
    }

    @Override
    public void booleanColumn(Column column) {
        if (pageReader.isNull(column)) {
            return;
        }
        boolean value = pageReader.getBoolean(column);
        Cell cell = getCell(column);
        cell.setCellValue(value);
    }

    @Override
    public void longColumn(Column column) {
        if (pageReader.isNull(column)) {
            return;
        }
        long value = pageReader.getLong(column);
        Cell cell = getCell(column);
        cell.setCellValue(value);
    }

    @Override
    public void doubleColumn(Column column) {
        if (pageReader.isNull(column)) {
            return;
        }
        double value = pageReader.getDouble(column);
        Cell cell = getCell(column);
        cell.setCellValue(value);
    }

    @Override
    public void stringColumn(Column column) {
        if (pageReader.isNull(column)) {
            return;
        }
        String value = pageReader.getString(column);
        Cell cell = getCell(column);
        cell.setCellValue(value);
    }

    @Override
    public void timestampColumn(Column column) {
        if (pageReader.isNull(column)) {
            return;
        }
        Instant timestamp = pageReader.getTimestampInstant(column);
        Date value = Date.from(timestamp);
        Cell cell = getCell(column);
        cell.setCellValue(value);
    }

    @Override
    public void jsonColumn(Column column) {
        if (pageReader.isNull(column)) {
            return;
        }
        JsonValue json = pageReader.getJsonValue(column);
        String value = json.toJson();
        Cell cell = getCell(column);
        cell.setCellValue(value);
    }

    protected Cell getCell(Column column) {
        Cell cell = CellUtil.getCell(getRow(), column.getIndex());

        ColumnOption option = getColumnOption(column);
        if (option != null) {
            Optional<String> formatOption = option.getDataFormat();
            if (formatOption.isPresent()) {
                String formatString = formatOption.get();
                CellStyle style = styleMap.get(formatString);
                if (style == null) {
                    Workbook book = sheet.getWorkbook();
                    style = book.createCellStyle();
                    CreationHelper helper = book.getCreationHelper();
                    short fmt = helper.createDataFormat().getFormat(formatString);
                    style.setDataFormat(fmt);
                    styleMap.put(formatString, style);
                }
                cell.setCellStyle(style);
            }
        }

        return cell;
    }

    protected final ColumnOption getColumnOption(Column column) {
        Map<String, ColumnOption> map = task.getColumnOptions();
        return map.get(column.getName());
    }

    private Map<String, CellStyle> styleMap = new HashMap<>();

    private Row getRow() {
        if (currentRow == null) {
            currentRow = sheet.createRow(rowIndex);
        }
        return currentRow;
    }

    public void endRecord() {
        rowIndex++;
        currentRow = null;
    }
}
