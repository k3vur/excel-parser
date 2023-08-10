package me.kevur.excelparser

data class ExcelParserConfig(
    val startRowIndex: Int,
    val sheet: SheetReference,
    val columnMapping: ColumnMapping
) {

    constructor(startRowIndex: Int, sheetName: String?, columnMapping: ColumnMapping) :
        this(startRowIndex, sheetName?.let { NameSheetReference(it) } ?: DefaultSheetReference, columnMapping)

    companion object {
        @JvmStatic
        fun of(startRowIndex: Int, sheet: SheetReference, columnMapping: ColumnMapping) =
            ExcelParserConfig(startRowIndex, sheet, columnMapping)

        @JvmStatic
        @JvmOverloads
        fun withDefaultSheet(startRowIndex: Int = 0, columnMapping: ColumnMapping) =
            ExcelParserConfig(startRowIndex, DefaultSheetReference, columnMapping)

        @JvmStatic
        @JvmOverloads
        fun withSheetName(startRowIndex: Int = 0, sheetName: String?, columnMapping: ColumnMapping): ExcelParserConfig =
            ExcelParserConfig(startRowIndex, sheetName?.let { NameSheetReference(it) } ?: DefaultSheetReference, columnMapping)

        @JvmStatic
        @JvmOverloads
        fun withSheetIndex(startRowIndex: Int = 0, sheetIndex: Int, columnMapping: ColumnMapping) =
            ExcelParserConfig(startRowIndex, IndexSheetReference(sheetIndex), columnMapping)
    }

}



sealed class SheetReference {
    companion object {
        @JvmStatic fun index(index: Int) = IndexSheetReference(index)
        @JvmStatic fun name(name: String) = NameSheetReference(name)
        @JvmField val DEFAULT: DefaultSheetReference = DefaultSheetReference
    }
}

data class NameSheetReference(val name: String) : SheetReference() {
    override fun toString(): String {
        return name
    }
}

data class IndexSheetReference(val index: Int) : SheetReference() {
    override fun toString(): String {
        return index.toString()
    }
}

object DefaultSheetReference : SheetReference()

