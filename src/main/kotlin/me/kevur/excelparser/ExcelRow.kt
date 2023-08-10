package me.kevur.excelparser

class ExcelRow(
    val index: Int,
    private val data: Map<String, String>
) {
    operator fun get(key: String): ExcelValue {
        return data[key]?.let(::ExcelValue) ?: ExcelValue(null)
    }

    fun isEmpty(): Boolean = data.isEmpty()

    val keys: Set<String> get() = this.data.keys

    fun forEach(action: (Map.Entry<String, ExcelValue>) -> Unit) =
        data.mapValues { (_, value) -> ExcelValue(value) }
            .forEach(action)
}
