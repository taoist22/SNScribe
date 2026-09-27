package com.sndocx

import java.io.File
import java.io.FileOutputStream
import java.util.zip.ZipEntry
import java.util.zip.ZipOutputStream

/**
 * A new, empty Word document: one empty paragraph, US Letter with 1" margins, and the
 * styles the editor's buttons use — Normal (the default, Calibri 11 pt), Heading 1,
 * Heading 2 and Title — so formatting a new document needs no style to be added later.
 * No theme, numbering or settings parts: Word does not need them, and the editor adds
 * numbering.xml when the first list is made.
 */
object DocxBlank {
    private const val W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    private const val R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

    private val PARTS = linkedMapOf(
        "[Content_Types].xml" to """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>""",
        "_rels/.rels" to """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
        "word/_rels/document.xml.rels" to """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>""",
        "word/document.xml" to """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="$W" xmlns:r="$R"><w:body><w:p/><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>""",
        "word/styles.xml" to """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="$W"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="200" w:after="80"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="240"/></w:pPr><w:rPr><w:sz w:val="56"/><w:szCs w:val="56"/></w:rPr></w:style></w:styles>""",
    )

    fun write(dest: File) {
        ZipOutputStream(FileOutputStream(dest)).use { out ->
            for ((name, xml) in PARTS) {
                out.putNextEntry(ZipEntry(name))
                out.write(xml.toByteArray(Charsets.UTF_8))
                out.closeEntry()
            }
        }
    }

    private const val TEMPLATE_MAIN = "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml"
    private const val DOCUMENT_MAIN = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"

    /**
     * Copies [template] (.docx or .dotx) to [dest] as a document: every part byte for byte,
     * except that a template's main part is declared a document in [Content_Types].xml.
     * Macro-enabled templates are refused.
     */
    fun fromTemplate(template: File, dest: File) {
        require(template.length() <= DocxReader.MAX_DOCX_BYTES) { "too large: ${template.length()} bytes" }
        java.util.zip.ZipFile(template).use { zip ->
            ZipOutputStream(FileOutputStream(dest)).use { out ->
                for (entry in zip.entries()) {
                    DocxReader.checkEntryName(entry.name)
                    var bytes = DocxReader.readEntry(zip, entry)
                    if (entry.name == "[Content_Types].xml") {
                        val xml = bytes.toString(Charsets.UTF_8)
                        require(!xml.contains("macroEnabled", ignoreCase = true)) { "Templates with macros can't be used" }
                        bytes = xml.replace(TEMPLATE_MAIN, DOCUMENT_MAIN).toByteArray(Charsets.UTF_8)
                    }
                    out.putNextEntry(ZipEntry(entry.name))
                    out.write(bytes)
                    out.closeEntry()
                }
            }
        }
    }

    /** `<name>.docx`, or `<name> 2.docx` … in [dir], never an existing file. */
    fun freeName(dir: File, name: String): File {
        val stem = name.trim().replace(Regex("[\\\\/:*?\"<>|]"), "-").removeSuffix(".docx").ifEmpty { "Untitled" }
        var file = File(dir, "$stem.docx")
        var n = 2
        while (file.exists()) file = File(dir, "$stem ${n++}.docx")
        return file
    }
}
