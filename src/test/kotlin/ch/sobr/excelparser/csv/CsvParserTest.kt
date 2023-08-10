package me.kevur.excelparser.csv

import me.kevur.excelparser.ExcelParser
import me.kevur.excelparser.TestUtils
import org.junit.jupiter.api.Test

class CsvParserTest {

    private val ENGLISH_UTF8_FILE = "test-english-utf-8.csv"
    private val GERMAN_ISO_8859_FILE = "test-german-iso-8859.csv"

    @Test
    fun `parse test english utf-8 csv as input stream`() {
        // given
        val testTableVerifier = TestUtils.TestTableVerifier()
        val inputStream =
            TestUtils.loadResourceFileAsInputStream(ENGLISH_UTF8_FILE)
        val csvOptions = CsvOptions.ENGLISH_UTF_8

        // when
        inputStream.use {
            ExcelParser.parseCsv(
                csvOptions,
                testTableVerifier.excelParserConfig,
                inputStream,
                testTableVerifier.rowHandler
            )
        }

        // then
        testTableVerifier.verify()
    }

    @Test
    fun `parse test english utf-8 csv as path`() {
        // given
        val testTableVerifier = TestUtils.TestTableVerifier()
        val path = TestUtils.loadResourceFileAsPath(ENGLISH_UTF8_FILE)
        val csvOptions = CsvOptions.ENGLISH_UTF_8

        // when
        ExcelParser.parseCsv(
            csvOptions,
            testTableVerifier.excelParserConfig,
            path,
            testTableVerifier.rowHandler
        )

        // then
        testTableVerifier.verify()
    }

    @Test
    fun `parse test german iso-8869 csv as input stream`() {
        // given
        val testTableVerifier = TestUtils.TestTableVerifier()
        val inputStream =
            TestUtils.loadResourceFileAsInputStream(GERMAN_ISO_8859_FILE)
        val csvOptions = CsvOptions.GERMAN_ISO_8859_1

        // when
        inputStream.use {
            ExcelParser.parseCsv(
                csvOptions,
                testTableVerifier.excelParserConfig,
                inputStream,
                testTableVerifier.rowHandler
            )
        }

        // then
        testTableVerifier.verify()
    }

    @Test
    fun `parse test english iso-8869 csv as path`() {
        // given
        val testTableVerifier = TestUtils.TestTableVerifier()
        val path = TestUtils.loadResourceFileAsPath(GERMAN_ISO_8859_FILE)
        val csvOptions = CsvOptions.GERMAN_ISO_8859_1

        // when
        ExcelParser.parseCsv(
            csvOptions,
            testTableVerifier.excelParserConfig,
            path,
            testTableVerifier.rowHandler
        )

        // then
        testTableVerifier.verify()
    }

}
