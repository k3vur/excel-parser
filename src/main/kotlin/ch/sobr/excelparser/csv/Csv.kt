package me.kevur.excelparser.csv

import me.kevur.excelparser.*
import me.kevur.excelparser.util.columnLettersToIndex
import org.apache.commons.csv.CSVFormat
import org.apache.commons.csv.CSVParser
import org.apache.commons.csv.CSVRecord
import org.apache.commons.io.input.BOMInputStream
import java.io.InputStream
import java.nio.charset.Charset
import java.nio.charset.StandardCharsets
import java.nio.file.Files
import java.nio.file.Path

internal class Csv(private val csvOptions: CsvOptions) : SpreadsheetFormat {

    override fun open(inputStream: InputStream): SpreadsheetFile {
        val reader = BOMInputStream(inputStream).bufferedReader(csvOptions.charset)
        val parser = csvOptions.commonsCsvFormat.parse(reader)
        return CsvSpreadsheet(parser)
    }

    override fun open(file: Path): SpreadsheetFile {
        val reader = BOMInputStream(Files.newInputStream(file)).bufferedReader(csvOptions.charset)
        val parser = csvOptions.commonsCsvFormat.parse(reader)
        return CsvSpreadsheet(parser)
    }

}

data class CsvOptions(
    val charset: Charset,
    val commonsCsvFormat: CSVFormat
) {
    companion object {
        @JvmField
        val ENGLISH_UTF_8 = CsvOptions(StandardCharsets.UTF_8, CSVFormat.EXCEL)

        @JvmField
        val GERMAN_ISO_8859_1 = CsvOptions(StandardCharsets.ISO_8859_1, CSVFormat.Builder.create(CSVFormat.EXCEL).setDelimiter(';').build())
    }
}

internal class CsvSpreadsheet(
    private val csvParser: CSVParser
) : SpreadsheetFile {
    override fun getSheet(name: String): Sheet = CsvSheet(csvParser)
    override fun getSheet(index: Int): Sheet = CsvSheet(csvParser)
    override fun getDefaultSheet(): Sheet = CsvSheet(csvParser)
    override fun close() = csvParser.close()
}

internal class CsvSheet(
    private val csvParser: CSVParser
) : Sheet {

    override fun process(config: ExcelParserConfig, rowHandler: RowHandler) {
        val processor = CsvProcessor(config, rowHandler)
        processor.process(csvParser)
    }

}

internal  class CsvProcessor(
    config: ExcelParserConfig,
    private val rowHandler: RowHandler
) {
    private val columnLookup = config.columnMapping.mapValues { (_, columnLetters) -> columnLettersToIndex(columnLetters) }
    private val startIndex = config.startRowIndex

    fun process(parser: CSVParser) = parser.asSequence()
        .filter { record -> record.recordIndex >= startIndex }
        .map { record ->
            ExcelRow(
                record.recordIndex,
                columnLookup
                    .mapValues { (_, columnIndex) ->
                        if (record.size() > columnIndex) record[columnIndex].trim()
                        else ""
                    }
                    .filter { (_, value) -> value.isNotBlank() }
            )
        }
        .filter { excelRow -> !excelRow.isEmpty() }
        .forEach { rowHandler(it) }

}

private val CSVRecord.recordIndex: Int get() = (this.recordNumber - 1).toInt()
