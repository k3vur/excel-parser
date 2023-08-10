package me.kevur.excelparser.xlsx

import me.kevur.excelparser.*
import me.kevur.excelparser.util.toLocalDateTime
import org.apache.poi.openxml4j.opc.OPCPackage
import org.apache.poi.openxml4j.opc.PackageAccess
import org.apache.poi.ss.usermodel.DataFormatter
import org.apache.poi.ss.usermodel.DateUtil
import org.apache.poi.util.XMLHelper
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable
import org.apache.poi.xssf.eventusermodel.XSSFReader
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler
import org.apache.poi.xssf.model.SharedStrings
import org.apache.poi.xssf.model.StylesTable
import org.xml.sax.InputSource
import java.io.InputStream
import java.nio.file.Path
import java.time.format.DateTimeFormatter
import javax.xml.parsers.ParserConfigurationException

internal object Xlsx : SpreadsheetFormat {

    private fun getSpreadsheet(opcPackage: OPCPackage): SpreadsheetFile {
        val xssfReader = XSSFReader(opcPackage)
        return XlsxSpreadsheet(opcPackage, xssfReader)
    }

    override fun open(inputStream: InputStream): SpreadsheetFile =
        OPCPackage.open(inputStream).use { getSpreadsheet(it) }

    override fun open(file: Path): SpreadsheetFile {
        val opcPackage = OPCPackage.open(file.toFile(), PackageAccess.READ)
        return getSpreadsheet(opcPackage)
    }

}

internal class XlsxSpreadsheet(
    private val opcPackage: OPCPackage,
    private val xssfReader: XSSFReader
) : SpreadsheetFile {

    override fun getSheet(name: String): Sheet? {
        val sheetIterator = (xssfReader.sheetsData as? XSSFReader.SheetIterator) ?: return null

        while (sheetIterator.hasNext()) {
            val currentSheet = sheetIterator.next()
            val sheetName = sheetIterator.sheetName
            if (sheetName == name) {
                return createSheet(currentSheet)
            }
        }

        return null
    }

    override fun getSheet(index: Int): Sheet? {
        val sheetIterator = xssfReader.sheetsData
        return sheetIterator.asSequence()
            .drop(index)
            .firstOrNull()
            ?.let(this::createSheet)
    }

    override fun getDefaultSheet(): Sheet {
        val data = xssfReader.sheetsData
        if (!data.hasNext()) {
            throw IllegalStateException("xlsx does not contain a sheet")
        }
        return createSheet(data.next())
    }

    private fun createSheet(inputStream: InputStream): XlsxSheet {
        val strings = ReadOnlySharedStringsTable(opcPackage)
        val styles = xssfReader.stylesTable
        return XlsxSheet(strings, styles, inputStream)
    }

    override fun close() {
        opcPackage.close()
    }

}

internal class XlsxSheet(
    private val strings: SharedStrings,
    private val stylesTable: StylesTable,
    private val inputStream: InputStream
) : Sheet {

    override fun process(config: ExcelParserConfig, rowHandler: RowHandler) {
        val sheetHandler = RowHandlerToXssfSheetContentHandlerAdapter(config, rowHandler)
        val sheetSource = InputSource(inputStream)
        try {
            with (XMLHelper.newXMLReader()) {
                contentHandler = XSSFSheetXMLHandler(
                    stylesTable, null, strings, sheetHandler, CustomDataFormatter, false
                )
                parse(sheetSource)
            }
        } catch (e: ParserConfigurationException) {
            throw ExcelParserException("SAX parser appears to be broken", e)
        }
    }

}

private object CustomDataFormatter : DataFormatter() {
    override fun formatRawCellContents(
        value: Double,
        formatIndex: Int,
        formatString: String?,
        use1904Windowing: Boolean
    ): String {
        return if (DateUtil.isADateFormat(formatIndex, formatString)) {
            DateUtil.getJavaDate(value, use1904Windowing).toLocalDateTime().format(DateTimeFormatter.ISO_DATE_TIME)
        } else {
            super.formatRawCellContents(value, formatIndex, formatString, use1904Windowing)
        }
    }
}

