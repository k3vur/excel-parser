package me.kevur.excelparser.xlsx

import me.kevur.excelparser.ExcelParserConfig
import me.kevur.excelparser.ExcelRow
import me.kevur.excelparser.RowHandler
import me.kevur.excelparser.util.ExcelUtil
import me.kevur.excelparser.util.invert
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler
import org.apache.poi.xssf.usermodel.XSSFComment

internal class RowHandlerToXssfSheetContentHandlerAdapter(
    config: ExcelParserConfig,
    private val rowHandler: RowHandler
) : XSSFSheetXMLHandler.SheetContentsHandler {

    private val columnLookup = config.columnMapping.invert()
    private val startIndex = config.startRowIndex
    private val currentRow: MutableMap<String, String> = mutableMapOf()

    override fun endRow(rowNum: Int) {
        if (rowNum >= this.startIndex && !currentRowIsEmpty()) {
            rowHandler(ExcelRow(rowNum, currentRow.toMap()))
        }
    }

    private fun currentRowIsEmpty() = currentRow.filter { (_, value) -> value.isNotBlank() }.isEmpty()

    override fun startRow(rowNum: Int) {
        if (rowNum >= this.startIndex) {
            currentRow.clear()
        }
    }

    override fun cell(cellReference: String?, formattedValue: String?, comment: XSSFComment?) {
        cellReference
            ?.let { getColumnName(it) }
            ?.let { key -> currentRow[key] = formattedValue?.trim() ?: "" }
    }

    private fun getColumnName(cellReference: String): String? {
        val letter = ExcelUtil.getColumnLetter(cellReference)
        return columnLookup[letter]
    }

}
