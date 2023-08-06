# Apache POI Excel formatter plugin for Embulk

Formats Excel files(xls, xlsx) for other file output plugins.  
This plugin uses Apache POI.

## Overview

* **Plugin type**: formatter
* Embulk 0.10 or later

## Configuration

* **spread_sheet_version**: Excel file version. `EXCEL97` or `EXCEL2007`. (string, default: `EXCEL2007`)
* **sheet_name**: sheet name. (string, default: `Sheet1`)
* **column_options**: see bellow. (hash, default: `{}`)

### column_options

* **data_format**: data format of Cell. (string, default: `null`)

## Example

```yaml
exec:
  min_output_tasks: 1	# output to one file

in:
  type: any input plugin type
...
    columns:
    - {name: time,     type: timestamp}
    - {name: purchase, type: timestamp}

out:
  type: file	# any file output plugin type
  path_prefix: /tmp/embulk-example/excel-out/sample_
  file_ext: xlsx
  formatter:
    type: poi_excel
    spread_sheet_version: EXCEL2007
    sheet_name: Sheet1
    column_options:
      time:     {data_format: "yyyy/mm/dd hh:mm:ss"}
      purchase: {data_format: "yyyy/mm/dd"}
```

### Note

The file name, file split or data order are decided by input/output plugin.  
If you'd like to process data and output Excel format, I think it's also one way to use [Asakusa Framework](http://www.asakusafw.com/) ([Excel Exporter](http://www.ne.jp/asahi/hishidama/home/tech/asakusafw/directio/excelformat.html>)).


## Install

1. download pom
   ```
   $ curl https://repo1.maven.org/maven2/io/github/hishidama/embulk/embulk-formatter-excel-poi/0.2.0/embulk-formatter-excel-poi-0.2.0.pom > embulk-formatter-excel-poi-0.2.0.pom
   ```

2. install dependencies
   ```
   $ mvn install -f embulk-formatter-excel-poi-0.2.0.pom
   ```

3. download and install jar
   ```
   $ export M2_REPO=$HOME/.m2/repository
   $ curl https://repo1.maven.org/maven2/io/github/hishidama/embulk/embulk-formatter-excel-poi/0.2.0/embulk-formatter-excel-poi-0.2.0.jar > $M2_REPO/io/github/hishidama/embulk/embulk-formatter-excel-poi/0.2.0/embulk-formatter-excel-poi-0.2.0.jar
   ```

4. add setting to $HOME/.embulk/embulk.properties
   ```
   plugins.formatter.poi_excel=maven:io.github.hishidama.embulk:excel-poi:0.2.0
   ```


## Build

```
$ ./gradlew package
```

### Build to local Maven repository

```
./gradlew generatePomFileForMavenJavaPublication
mvn install -f build/publications/mavenJava/pom-default.xml
./gradlew publishToMavenLocal
```

