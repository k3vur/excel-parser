package me.kevur.excelparser

import org.junit.jupiter.api.Assertions.assertEquals
import org.junit.jupiter.api.fail
import java.io.InputStream
import java.nio.file.Path
import java.nio.file.Paths

object TestUtils {

    fun loadResourceFileAsPath(name: String): Path {
        val url = TestUtils::class.java.classLoader.getResource(name) ?: fail("cannot load test file $name")
        return Paths.get(url.toURI())
    }

    fun loadResourceFileAsInputStream(name: String): InputStream {
        return TestUtils::class.java.classLoader.getResourceAsStream(name) ?: fail("cannot load test file $name")
    }

    val testFileColumnMapping = mapOf(
        "prop1" to "A",
        "prop2" to "B",
        "prop3" to "C",
        "prop4" to "D"
    )

    internal data class ExpectedRowData(
        val rowIndex: Int,
        val rowData: Map<String, String>
    )

    class TestTableVerifier {

        private val expectedData = listOf(
            ExpectedRowData(0, mapOf(
                "prop1" to "1a",
                "prop2" to "2a",
                "prop3" to "3a",
                "prop4" to "4a"
            )),
            ExpectedRowData(2, mapOf(
                "prop1" to "3a",
                "prop2" to "3b",
                "prop3" to "3c",
                "prop4" to "3d"
            )),
            ExpectedRowData(3, mapOf(
                "prop1" to "4a",
                "prop4" to "4d"
            )),
            ExpectedRowData(4, mapOf(
                "prop2" to "5b",
                "prop3" to "5c",
                "prop4" to "5d"
            )),
            ExpectedRowData(5, mapOf(
                "prop1" to "6a",
                "prop2" to "6b"
            ))
        )

        val excelParserConfig = ExcelParserConfig(0, null, testFileColumnMapping)
        private val rows = mutableListOf<ExcelRow>()
        val rowHandler: RowHandler = { row -> rows.add(row) }

        fun verify() {
            assertEquals(5, rows.size, "row handler should have been called 5 times")
            expectedData.zip(rows).forEach { (expectedRowData, actualRow) ->
                assertEquals(actualRow.index, expectedRowData.rowIndex, "row index does not match")
                assertEquals(actualRow.keys, expectedRowData.rowData.keys, "keys in row do not match")
                expectedRowData.rowData.forEach { (key, value) ->
                    assertEquals(actualRow[key].rawValue, value, "cell value does not match")
                }
            }
        }
    }

}
