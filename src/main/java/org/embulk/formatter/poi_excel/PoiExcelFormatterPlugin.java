package org.embulk.formatter.poi_excel;

import java.io.IOException;
import java.io.UncheckedIOException;
import java.text.MessageFormat;
import java.util.Map;
import java.util.Optional;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.embulk.config.ConfigSource;
import org.embulk.config.TaskSource;
import org.embulk.formatter.poi_excel.visitor.PoiExcelColumnVisitor;
import org.embulk.spi.Exec;
import org.embulk.spi.FileOutput;
import org.embulk.spi.FormatterPlugin;
import org.embulk.spi.Page;
import org.embulk.spi.PageOutput;
import org.embulk.spi.PageReader;
import org.embulk.spi.Schema;
import org.embulk.util.config.Config;
import org.embulk.util.config.ConfigDefault;
import org.embulk.util.config.ConfigMapper;
import org.embulk.util.config.ConfigMapperFactory;
import org.embulk.util.config.Task;
import org.embulk.util.config.TaskMapper;
import org.embulk.util.file.FileOutputOutputStream;
import org.embulk.util.file.FileOutputOutputStream.CloseMode;

public class PoiExcelFormatterPlugin implements FormatterPlugin {
//  private final Logger log = LoggerFactory.getLogger(getClass());

    public static final String TYPE = "poi_excel";

    public interface PluginTask extends Task, TimestampFormatterTask {
        @Config("spread_sheet_version")
        @ConfigDefault("\"EXCEL2007\"")
        public SpreadsheetVersion getSpreadsheetVersion();

        @Config("sheet_name")
        @ConfigDefault("\"Sheet1\"")
        public String getSheetName();

        @Config("column_options")
        @ConfigDefault("{}")
        public Map<String, ColumnOption> getColumnOptions();
    }

    // From org.embulk.spi.time.TimestampFormatter.Task
    public interface TimestampFormatterTask {
        @Config("default_timezone")
        @ConfigDefault("\"UTC\"")
        public String getDefaultTimeZoneId();

        @Config("default_timestamp_format")
        @ConfigDefault("\"%Y-%m-%d %H:%M:%S.%6N %z\"")
        public String getDefaultTimestampFormat();
    }

    public interface ColumnOption extends Task, TimestampColumnOption {
        @Config("data_format")
        @ConfigDefault("null")
        public Optional<String> getDataFormat();
    }

    // org.embulk.spi.time.TimestampFormatter.TimestampColumnOption
    public interface TimestampColumnOption {
        @Config("timezone")
        @ConfigDefault("null")
        public Optional<String> getTimeZoneId();

        @Config("format")
        @ConfigDefault("null")
        public Optional<String> getFormat();
    }

    protected static final ConfigMapper CONFIG_MAPPER;
    protected static final TaskMapper TASK_MAPPER;
    static {
        ConfigMapperFactory factory = ConfigMapperFactory.builder().addDefaultModules().build();
        CONFIG_MAPPER = factory.createConfigMapper();
        TASK_MAPPER = factory.createTaskMapper();
    }

    @Override
    public void transaction(ConfigSource config, Schema schema, FormatterPlugin.Control control) {
        PluginTask task = CONFIG_MAPPER.map(config, PluginTask.class);

        control.run(task.toTaskSource());
    }

    @Override
    public PageOutput open(TaskSource taskSource, final Schema schema, FileOutput output) {
        final PluginTask task = TASK_MAPPER.map(taskSource, PluginTask.class);

        final Sheet sheet = newWorkbook(task);

        final FileOutputOutputStream stream = new FileOutputOutputStream(output, Exec.getBufferAllocator(), CloseMode.CLOSE);
        stream.nextFile();

        return new PageOutput() {
            private final PageReader pageReader = Exec.getPageReader(schema);

            @Override
            public void add(Page page) {
                pageReader.setPage(page);
                PoiExcelColumnVisitor visitor = new PoiExcelColumnVisitor(task, sheet, pageReader);
                while (pageReader.nextRecord()) {
                    schema.visitColumns(visitor);
                    visitor.endRecord();
                }
            }

            @Override
            public void finish() {
                Workbook book = sheet.getWorkbook();
                try (FileOutputOutputStream os = stream) {
                    book.write(os);
                    os.finish();
                } catch (IOException e) {
                    throw new UncheckedIOException(e.getMessage(), e);
                }
            }

            @Override
            public void close() {
                stream.close();
            }
        };
    }

    @SuppressWarnings("resource")
    protected Sheet newWorkbook(PluginTask task) {
        Workbook book;
        {
            SpreadsheetVersion version = task.getSpreadsheetVersion();
            switch (version) {
            case EXCEL97:
                book = new HSSFWorkbook();
                break;
            case EXCEL2007:
                book = new XSSFWorkbook();
                break;
            default:
                throw new UnsupportedOperationException(MessageFormat.format("unsupported spread_sheet_version={0}", version));
            }
        }

        String sheetName = task.getSheetName();
        return book.createSheet(sheetName);
    }
}
