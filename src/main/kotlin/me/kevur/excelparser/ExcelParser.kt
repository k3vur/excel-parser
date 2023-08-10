package me.kevur.excelparser

import me.kevur.excelparser.csv.Csv
import me.kevur.excelparser.csv.CsvOptions
import me.kevur.excelparser.xlsx.Xlsx
import java.io.InputStream
import java.nio.file.Path
import java.util.function.Consumer

typealias RowHandler = (row: ExcelRow) -> Unit
typealias ColumnMapping = Map<String, String>

object ExcelParser {

    private fun parse(config: ExcelParserConfig, format: SpreadsheetFormat, inputStream: InputStream, rowHandler: RowHandler) {
        val worksheet = format.open(inputStream)
        parseWorksheet(config, worksheet, rowHandler)
    }

    private fun parse(config: ExcelParserConfig, format: SpreadsheetFormat, file: Path, rowHandler: RowHandler) {
        val worksheet = format.open(file)
        parseWorksheet(config, worksheet, rowHandler)
    }

    fun parseXlsx(config: ExcelParserConfig, inputStream: InputStream, rowHandler: RowHandler) =
        parse(config, Xlsx, inputStream, rowHandler)

    fun parseXlsx(config: ExcelParserConfig, file: Path, rowHandler: RowHandler) = parse(config, Xlsx, file, rowHandler)

    @JvmStatic
    fun parseXlsx(config: ExcelParserConfig, inputStream: InputStream, rowHandler: Consumer<ExcelRow>) =
        parseXlsx(config, inputStream, rowHandler::accept)

    @JvmStatic
    fun parseXlsx(config: ExcelParserConfig, file: Path, rowHandler: Consumer<ExcelRow>) =
        parseXlsx(config, file, rowHandler::accept)

    fun parseCsv(csvOptions: CsvOptions, config: ExcelParserConfig, file: Path, rowHandler: RowHandler) =
        parse(config, Csv(csvOptions), file, rowHandler)

    fun parseCsv(csvOptions: CsvOptions, config: ExcelParserConfig, inputStream: InputStream, rowHandler: RowHandler) =
        parse(config, Csv(csvOptions), inputStream, rowHandler)

    @JvmStatic
    fun parseCsv(csvOptions: CsvOptions, config: ExcelParserConfig, file: Path, rowHandler: Consumer<ExcelRow>) =
        parseCsv(csvOptions, config, file, rowHandler::accept)

    @JvmStatic
    fun parseCsv(csvOptions: CsvOptions, config: ExcelParserConfig, inputStream: InputStream, rowHandler: Consumer<ExcelRow>) =
        parseCsv(csvOptions, config, inputStream, rowHandler::accept)

    private fun parseWorksheet(config: ExcelParserConfig, spreadsheet: SpreadsheetFile, rowHandler: RowHandler) {
        val sheet = spreadsheet.getSheet(config.sheet)
            ?: throw ExcelParserException("could not find sheet \"${config.sheet}\"")
        sheet.process(config, rowHandler)
        spreadsheet.close()
    }

}

internal interface SpreadsheetFormat {
    fun open(inputStream: InputStream): SpreadsheetFile
    fun open(file: Path): SpreadsheetFile
}

internal interface SpreadsheetFile {
    fun getSheet(name: String): Sheet?
    fun getSheet(index: Int): Sheet?
    fun getDefaultSheet(): Sheet
    fun close()

    fun getSheet(sheetReference: SheetReference): Sheet? = when (sheetReference) {
        is NameSheetReference -> getSheet(sheetReference.name)
        is IndexSheetReference -> getSheet(sheetReference.index)
        is DefaultSheetReference -> getDefaultSheet()
    }
}

internal interface Sheet {
    fun process(config: ExcelParserConfig, rowHandler: RowHandler)
}
