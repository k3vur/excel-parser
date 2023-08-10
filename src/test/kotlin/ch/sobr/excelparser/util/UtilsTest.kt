package me.kevur.excelparser.util

import org.junit.jupiter.api.Assertions.assertEquals
import org.junit.jupiter.api.Test

class UtilsTest {

    val testTable = listOf(
        "A" to 0,
        "B" to 1,
        "Z" to 25,
        "AA" to 26,
        "AB" to 27,
        "AZ" to 51,
        "BA" to 52,
        "BZ" to 77,
        "CA" to 78,
        "CZ" to 103,
        "ZA" to 676,
        "ZZ" to 701,
        "AAA" to 702
    )

    @Test
    fun `test column to index conversion`() {
        testTable.forEach {  (letter, index) ->
            assertEquals(index, columnLettersToIndex(letter))
        }
    }

}
