# Excel Parser

A small Excel library for JVM languages to parse and process Excel and CSV files line by line.

* Support XLSX Files and CSV files of arbitrary size (no 65k rows limit)
* It's just a jar file to include and call
* No support for *.xls files
* No magic and hyper specific groovy / grails stuff

## Howto / Example

* add to your `build.gradle` file:
    ```gradle
    implementation 'me.kevur:excel-parser:0.5.0'
    ```

  
### XLSX

* Call parser as follows (example in groovy):


```groovy
    final ExcelParserConfig config = new ExcelParserConfig(1, null, [
        "customerNumber": "B",
        "name": "C",
        "routeToMarket": "A"
    ])
    def path = Paths.get("path/to/file")
    ExcelParser.parseXlsx(config, path) { row ->
        println(row.get("customerNumber").numberString)
        println(row.get("name").string)
        println(row.get("routeToMarket").nullableString)
    }
  
    // you can also use inputStream but in case of XLSX files, it will use much more RAM.
    // usually it's a good idea to cache on file system anyways (MultiPartFiles may be closed unexpectedly...)
    ExcelParser.parseXlsx(config, inputStream) { row -> /* process row */ }
```

  
### CSV

* Do the same as above, but call `ExcelParser.parseCsv()`
* Note that `parseCsv()` must receive a `CsvOptions` argument that specifies encoding and apache commons `CSVFormat` (commas vs. semicolons, ...)
* CSVs can't have multiple worksheets, so the worksheet argument in `ExcelParserConfig` is ignored


```groovy
    // example 1: standard english excel exported CSV
    final CsvOptions csvOptions = CsvOptions.ENGLISH_UTF_8
    // example 2: german excel exported CSV with semicolons and ISO-8859-1
    final CsvOptions csvOptions = CsvOptions.GERMAN_ISO_8859_1

  
    // call parser
    final ExcelParser.parseCsv(csvOptions, config, path) { row -> /* process rows */ }
  
    // in case of CSV, calling the inputStream pendant won't use more memory
    final ExcelParser.parseCsv(csvOptions, config, inputStream) { row -> /* process rows */ } 
```
    

## Good to know

* You can get a specific worksheet by providing it in the `ExcelParserConfig`, or just get the first one by declaring null
* The object you get as `row` is of class `ExcelRow`, the `.get()` method returns an `ExcelValue`
* If a row is completely empty, the `rowHandler` won't be called for it
* You ALWAYS get a `ExcelValue` on calling `ExcelRow.get()`. It can be empty though (check with `isEmpty()` property or use `nullableXXX` properties).
* `ExcelValue`s provide getters for different data types:
    * Strings (nullable / non-nullable)
    * integers, doubles, BigDecimals
    * booleans (nullable / non-nullable)
    * LocalDate and LocalDateTime
    * numberString (no funny scientific notifications, no trailing 0s and not decimal digits if not needed)
* There is a `forEach()` method on `ExcelRow` if you need to iterate
* There is a `getKeys()` method on `ExcelRow` if you need to check for existing columns in a row (columns might be empty?)
* if you need more data types:
    * implement here and push ;)
    * there is a `rawValue` property on `ExcelValue` that you can abuse for whatever magic you need
* CSV parser will correctly handle BOMs
