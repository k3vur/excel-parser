package me.kevur.excelparser.util

import org.apache.poi.ss.usermodel.DateUtil
import org.apache.poi.ss.util.CellReference
import java.time.LocalDate
import java.time.LocalDateTime
import java.time.ZoneId
import java.util.*

internal object ExcelUtil {
    fun getColumnLetter(cellReference: String): String = CellReference(cellReference).cellRefParts[2]

    fun getDate(excelDateStr: String): Date = DateUtil.getJavaDate(DateUtil.parseDateTime(excelDateStr))
}

internal fun Date.toLocalDate(): LocalDate = this.toInstant().atZone(ZoneId.systemDefault()).toLocalDate()
internal fun Date.toLocalDateTime(): LocalDateTime = this.toInstant().atZone(ZoneId.systemDefault()).toLocalDateTime()

internal fun <K,V> Map<K,V>.invert(): Map<V,K> {
    val invertedEntries = this.entries.map { (k, v) -> v to k }
    return mapOf(*invertedEntries.toTypedArray())
}

internal fun columnLettersToIndex(columnLetters: String): Int {
    val charArray = columnLetters.toCharArray()
    charArray.forEach { require(it.lowercaseChar() - 'a' in 0..25) { "char $it s not an allowed column letter"} }
    val columnNumber = charArray
        .map { it.lowercaseChar() - 'a' }
        .fold(0) { acc, i -> acc * 26 + i + 1 }
    return columnNumber - 1
}
