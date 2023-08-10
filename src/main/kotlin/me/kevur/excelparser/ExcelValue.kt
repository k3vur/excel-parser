package me.kevur.excelparser

import java.math.BigDecimal
import java.text.DecimalFormat
import java.text.NumberFormat
import java.text.ParseException
import java.time.LocalDate
import java.time.LocalDateTime
import java.time.format.DateTimeFormatter
import java.time.format.DateTimeParseException
import java.util.*

data class ExcelValue(
    val rawValue: String?
) {
    val isEmpty: Boolean get() = rawValue == null

    fun getNumber(format: NumberFormat): Number? =
        try { nullableString?.let(format::parse) } catch (e: NumberFormatException) { null }
    fun getOptionalNumber(format: NumberFormat): Optional<Number> = Optional.ofNullable(getNumber(format))

    val integer: Int? get() = nullableString?.toIntOrNull()
    fun getInteger(format: NumberFormat): Int? = getNumber(format)?.toInt()
    val optionalInteger: Optional<Int> get() = Optional.ofNullable(integer)
    fun getOptionalInteger(format: NumberFormat): Optional<Int> = Optional.ofNullable(getInteger(format))

    val double: Double? get() = nullableString?.toDoubleOrNull()
    fun getDouble(format: NumberFormat): Double? = getNumber(format)?.toDouble()
    val optionalDouble: Optional<Double> get() = Optional.ofNullable(double)
    fun getOptionalDouble(format: NumberFormat): Optional<Double> = Optional.ofNullable(getDouble(format))

    val bigDecimal: BigDecimal? get() = nullableString?.toBigDecimalOrNull()
    fun getBigDecimal(format: DecimalFormat): BigDecimal? = nullableString?.let {
        val bdFormat = (format.clone() as DecimalFormat).apply { isParseBigDecimal = true }
        return@let try { bdFormat.parse(it) as BigDecimal } catch (e: ParseException) { null }
    }
    val optionalBigDecimal: Optional<BigDecimal> get() = Optional.ofNullable(bigDecimal)
    fun getOptionalBigDecimal(format: DecimalFormat): Optional<BigDecimal> = Optional.ofNullable(getBigDecimal(format))

    fun getDate(formatter: DateTimeFormatter): LocalDate? = this.nullableString?.let {
        try {
            LocalDate.parse(it, formatter)
        } catch (e: DateTimeParseException) {
            null
        }
    }
    val date: LocalDate? get() = this.nullableString?.let { dateString ->
        val format =
            if (dateString.length == 10) DateTimeFormatter.ISO_DATE
            else DateTimeFormatter.ISO_DATE_TIME
        LocalDate.parse(dateString, format)
    }
    fun getOptionalDate(formatter: DateTimeFormatter): Optional<LocalDate> = Optional.ofNullable(getDate(formatter))
    val optionalDate: Optional<LocalDate> get() = Optional.ofNullable(date)

    val dateTime: LocalDateTime? get() = this.getDateTime(DateTimeFormatter.ISO_DATE_TIME)
    fun getDateTime(formatter: DateTimeFormatter) = this.nullableString?.let {
        try {
            LocalDateTime.parse(it, formatter)
        } catch (e: DateTimeParseException) {
            null
        }
    }
    val optionalDateTime: Optional<LocalDateTime> get() = Optional.ofNullable(dateTime)
    fun getOptionalDateTime(formatter: DateTimeFormatter): Optional<LocalDateTime> = Optional.ofNullable(getDateTime(formatter))


    /**
     * returns trimmed string value if exists, or empty string if not
     */
    val string: String get() = rawValue?.trim() ?: ""

    /**
     * returns trimmed string or null if the value is empty
     */
    val nullableString: String?
        get() = rawValue?.trim()?.let {
            if (it.isEmpty()) null
            else it
        }
    val optionalString: Optional<String> get() = Optional.ofNullable(nullableString)

    /**
     * returns a clean number string, meaning:
     * - if it's an int, return an int (without decimals)
     * - if it's a decimal return that, with excess 0s trimmed, and without scientific notation
     * - else return the string
     */
    val numberString: String?
        get() {
            val strValue = this.nullableString ?: return null
            val doubleValue = strValue.toDoubleOrNull()
            return when {
                doubleValue == null -> strValue
                doubleValue % 1 == 0.0 -> String.format("%.0f", doubleValue)
                else -> String.format("%f").replace(Regex("0+$"), "")
            }
        }
    val optionalNumberString: Optional<String> get() = Optional.ofNullable(numberString)

    private val defaultTrueValues = setOf("yes", "y", "1", "true")
    private val defaultFalseValues = setOf("no", "n", "0", "false")

    /**
     * returns true if the value is one of the provided allowedTrueValues, or false otherwise
     */
    @JvmOverloads
    fun getBoolean(allowedTrueValues: Iterable<String> = defaultTrueValues): Boolean {
        val lowercaseTrueValues = allowedTrueValues.map { it.lowercase() }.toSet()
        return this.nullableString?.lowercase() in lowercaseTrueValues
    }

    /**
     * returns true if the value is one of the provided allowedTrueValues,
     * false if it's one of the provided allowedFalseValues,
     * or null otherwise
     */
    @JvmOverloads
    fun getNullableBoolean(
        allowedTrueValues: Iterable<String> = defaultTrueValues,
        allowedFalseValues: Iterable<String> = defaultFalseValues
    ): Boolean? {
        val lowercaseTrueValues = allowedTrueValues.map(String::lowercase).toSet()
        val lowercaseFalseValues = allowedFalseValues.map(String::lowercase).toSet()
        return when (this.nullableString?.lowercase()) {
            in lowercaseTrueValues -> true
            in lowercaseFalseValues -> false
            else -> null
        }
    }

    @JvmOverloads
    fun getOptionalBoolean(
        allowedTrueValues: Iterable<String> = defaultTrueValues,
        allowedFalseValues: Iterable<String> = defaultFalseValues
    ): Optional<Boolean> = Optional.ofNullable(getNullableBoolean(allowedTrueValues, allowedFalseValues))
}
