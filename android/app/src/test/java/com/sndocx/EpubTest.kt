package com.sndocx

import org.junit.Assert.assertEquals
import org.junit.Test
import java.io.File
import java.nio.file.Files
import java.util.zip.ZipEntry
import java.util.zip.ZipOutputStream

class EpubTest {
    @Test
    fun readsTheEpubsOwnMetadata() {
        val dir = Files.createTempDirectory("epub").toFile()
        val epub = File(dir, "book.epub")
        ZipOutputStream(epub.outputStream()).use { z ->
            fun put(name: String, text: String) {
                z.putNextEntry(ZipEntry(name)); z.write(text.toByteArray()); z.closeEntry()
            }
            put("mimetype", "application/epub+zip")
            put("META-INF/container.xml", """<?xml version="1.0"?><container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container"><rootfiles><rootfile full-path="OEBPS/content.opf" media-type="application/oebps-package+xml"/></rootfiles></container>""")
            put("OEBPS/content.opf", """<?xml version="1.0"?><package xmlns="http://www.idpf.org/2007/opf" version="3.0"><metadata xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Financial Accounting</dc:title><dc:creator>Jane Q. Smith</dc:creator><dc:creator>Al Wu</dc:creator><dc:date>2021-05-01</dc:date><dc:publisher>OpenStax</dc:publisher></metadata></package>""")
        }
        val info = DocxReader.epubInfo(epub)
        assertEquals("Financial Accounting", info.title)
        assertEquals(listOf("Jane Q. Smith", "Al Wu"), info.creators)
        assertEquals("2021-05-01", info.date)
        assertEquals("OpenStax", info.publisher)
    }
}
