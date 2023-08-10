package me.kevur.excelparser.xlsx

import me.kevur.excelparser.*
import org.junit.jupiter.api.Test
import org.junit.jupiter.api.fail

class XlsxParserTest {

    @Test
    fun `parse test xlsx as InputStream`() {
        // given
        val testTableVerifier = TestUtils.TestTableVerifier()
        val testInputStream =
            TestUtils.loadResourceFileAsInputStream("test.xlsx")

        // when
        testInputStream.use { inputStream ->
            ExcelParser.parseXlsx(
                testTableVerifier.excelParserConfig,
                inputStream,
                testTableVerifier.rowHandler
            )
        }

        // then
        testTableVerifier.verify()
    }

    @Test
    fun `parse test xlsx as Path`() {
        // given
        val testTableVerifier = TestUtils.TestTableVerifier()
        val path = TestUtils.loadResourceFileAsPath("test.xlsx")

        // when
        ExcelParser.parseXlsx(
            testTableVerifier.excelParserConfig,
            path,
            testTableVerifier.rowHandler
        )

        // then
        testTableVerifier.verify()
    }

    @Test
    fun `ignore rows with only blank cells`() {
        val file = TestUtils.loadResourceFileAsPath("test-blank-space.xlsx")
        val config = ExcelParserConfig(0, null, columnMapping = mapOf(
            "col1" to "A",
            "col2" to "B"
        ))
        val rowHandler = { excelRow: ExcelRow -> if (excelRow.index == 1) fail("should not be called on 2nd row") }

        ExcelParser.parseXlsx(config, file, rowHandler)
    }

    @Test
    fun `get sheet by name`() {
        // given
        val testTableVerifier = TestUtils.TestTableVerifier()
        val file = TestUtils.loadResourceFileAsPath("test-multisheet.xlsx")
        val config = testTableVerifier.excelParserConfig.copy(sheet = SheetReference.name("namedsheet 2"))

        // when
        ExcelParser.parseXlsx(config, file, testTableVerifier.rowHandler)

        // then
        testTableVerifier.verify()
    }

    @Test
    fun `get sheet by index`() {
        // given
        val testTableVerifier = TestUtils.TestTableVerifier()
        val file = TestUtils.loadResourceFileAsPath("test-multisheet.xlsx")
        val config = testTableVerifier.excelParserConfig.copy(sheet = SheetReference.index(1))

        // when
        ExcelParser.parseXlsx(config, file, testTableVerifier.rowHandler)

        // then
        testTableVerifier.verify()
    }

}
