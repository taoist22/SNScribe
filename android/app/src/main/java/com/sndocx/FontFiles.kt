package com.sndocx

import java.io.File
import java.io.RandomAccessFile

/**
 * Font files on the device, without Android: reading a TrueType/OpenType font's family and
 * style from its 'name' table, and the file names a family is likely to have.
 *
 * Plugins cannot list a folder (FileUtils.listFiles is unavailable, and java.io listing of
 * user storage returns null), but they can test whether a named file exists. So the fonts a
 * document uses are looked for under guessed names in MyStyle/Fonts, and fonts the user
 * picks by hand are named from their own name table.
 */
object FontFiles {
    /** Where the device keeps user fonts (sn-datetime's custom fonts come from here). */
    val DIRS = listOf("/storage/emulated/0/MyStyle/Fonts", "/storage/emulated/0/MyStyle")

    data class Info(val family: String, val bold: Boolean, val italic: Boolean)

    /** Family and style from the font's name table; null when the file is not a readable font. */
    fun info(file: File): Info? = runCatching {
        RandomAccessFile(file, "r").use { f ->
            var base = 0L
            if (tag(f, 0) == "ttcf") base = u32(f, 12) // a collection: its first font
            val numTables = u16(f, base + 4)
            var nameOffset = -1L
            for (i in 0 until numTables) {
                val rec = base + 12 + i * 16L
                if (tag(f, rec) == "name") nameOffset = u32(f, rec + 8)
            }
            if (nameOffset < 0) return null
            val count = u16(f, nameOffset + 2)
            val strings = nameOffset + u16(f, nameOffset + 4)
            val names = HashMap<Int, String>()
            val rank = HashMap<Int, Int>()
            for (i in 0 until count) {
                val r = nameOffset + 6 + i * 12L
                val platform = u16(f, r)
                val encoding = u16(f, r + 2)
                val language = u16(f, r + 4)
                val id = u16(f, r + 6)
                val len = u16(f, r + 8)
                val off = u16(f, r + 10)
                if (id !in setOf(1, 2, 16, 17)) continue
                // Prefer Windows Unicode US English, then any Windows Unicode, then Mac Roman.
                val score = when {
                    platform == 3 && (encoding == 1 || encoding == 10) && language == 0x409 -> 3
                    platform == 3 && (encoding == 1 || encoding == 10) -> 2
                    platform == 1 && encoding == 0 -> 1
                    else -> 0
                }
                if (score == 0 || score <= (rank[id] ?: 0)) continue
                val bytes = ByteArray(len)
                f.seek(strings + off)
                f.readFully(bytes)
                names[id] = if (platform == 3) String(bytes, Charsets.UTF_16BE) else String(bytes, Charsets.ISO_8859_1)
                rank[id] = score
            }
            val family = (names[16] ?: names[1])?.trim()?.takeIf { it.isNotEmpty() } ?: return null
            val style = (names[17] ?: names[2] ?: "").lowercase()
            Info(family, "bold" in style || "black" in style || "heavy" in style, "italic" in style || "oblique" in style)
        }
    }.getOrNull()

    /** Style index as Android's Typeface uses it: 0 normal, 1 bold, 2 italic, 3 bold italic. */
    fun styleIndex(bold: Boolean, italic: Boolean) = (if (bold) 1 else 0) + (if (italic) 2 else 0)

    /** Windows short names of common Office fonts: regular, bold, italic, bold italic. */
    private val WINDOWS = mapOf(
        "arial" to listOf("arial", "arialbd", "ariali", "arialbi"),
        "calibri" to listOf("calibri", "calibrib", "calibrii", "calibriz"),
        "calibri light" to listOf("calibril", "calibrib", "calibrili", "calibriz"),
        "cambria" to listOf("cambria", "cambriab", "cambriai", "cambriaz"),
        "candara" to listOf("candara", "candarab", "candarai", "candaraz"),
        "century gothic" to listOf("gothic", "gothicb", "gothici", "gothicbi"),
        "comic sans ms" to listOf("comic", "comicbd", "comici", "comicz"),
        "consolas" to listOf("consola", "consolab", "consolai", "consolaz"),
        "constantia" to listOf("constan", "constanb", "constani", "constanz"),
        "corbel" to listOf("corbel", "corbelb", "corbeli", "corbelz"),
        "courier new" to listOf("cour", "courbd", "couri", "courbi"),
        "garamond" to listOf("gara", "garabd", "garait", "garabi"),
        "georgia" to listOf("georgia", "georgiab", "georgiai", "georgiaz"),
        "palatino linotype" to listOf("pala", "palab", "palai", "palabi"),
        "segoe ui" to listOf("segoeui", "segoeuib", "segoeuii", "segoeuiz"),
        "tahoma" to listOf("tahoma", "tahomabd", "tahoma", "tahomabd"),
        "times new roman" to listOf("times", "timesbd", "timesi", "timesbi"),
        "trebuchet ms" to listOf("trebuc", "trebucbd", "trebucit", "trebucbi"),
        "verdana" to listOf("verdana", "verdanab", "verdanai", "verdanaz"),
        "aptos" to listOf("aptos", "aptos-bold", "aptos-italic", "aptos-bolditalic"),
    )

    /**
     * File names worth trying for [family] in [style] (0–3), most likely first, without
     * folder: "Times New Roman Bold.ttf", "TimesNewRoman-Bold.ttf", "timesbd.ttf" …
     */
    fun candidates(family: String, style: Int): List<String> {
        val names = linkedSetOf<String>()
        val spaced = family.trim()
        val joined = spaced.replace(" ", "")
        val suffixes = when (style) {
            1 -> listOf(" Bold", "-Bold", "Bold", "_Bold", "bd", "b")
            2 -> listOf(" Italic", "-Italic", "Italic", "_Italic", "i", "it")
            3 -> listOf(" Bold Italic", "-BoldItalic", "BoldItalic", "_BoldItalic", "bi", "z")
            else -> listOf("", " Regular", "-Regular", "Regular", "_Regular")
        }
        for (stem in listOf(spaced, joined, spaced.lowercase(), joined.lowercase())) {
            for (sfx in suffixes) {
                names.add("$stem$sfx")
            }
        }
        WINDOWS[spaced.lowercase()]?.getOrNull(style)?.let { names.add(it); names.add(it.uppercase()) }
        return names.flatMap { listOf("$it.ttf", "$it.otf", "$it.TTF", "$it.OTF") }
    }

    /**
     * The file for [family] in [style] under [dirs], or null: by its usual file names first,
     * then by the family and style written inside the font files there (and one folder down,
     * as a download's "static" folder) — so a font is found however its file is named.
     */
    fun find(family: String, style: Int, dirs: List<String> = DIRS): File? {
        for (dir in dirs) {
            for (name in candidates(family, style)) {
                val f = File(dir, name)
                if (f.isFile) return f
            }
        }
        return index(dirs)[family.lowercase() to style]
    }

    private var indexed: Pair<List<String>, Map<Pair<String, Int>, File>>? = null

    /** family (lower case) and style → file, from the fonts' own name tables. Built once per session. */
    @Synchronized
    fun index(dirs: List<String> = DIRS): Map<Pair<String, Int>, File> {
        indexed?.takeIf { it.first == dirs }?.let { return it.second }
        val out = HashMap<Pair<String, Int>, File>()
        fun scan(dir: File, depth: Int) {
            val files = dir.listFiles() ?: return
            for (f in files.sortedBy { it.name }) {
                if (f.isDirectory) {
                    if (depth > 0) scan(f, depth - 1)
                    continue
                }
                if (!f.name.lowercase().let { it.endsWith(".ttf") || it.endsWith(".otf") }) continue
                // A variable font file stands for its regular (or italic) style.
                val i = info(f) ?: continue
                out.putIfAbsent(i.family.lowercase() to styleIndex(i.bold, i.italic), f)
            }
        }
        for (d in dirs) scan(File(d), 1)
        indexed = dirs to out
        return out
    }

    /** Forget the index (fonts were added). */
    fun reindex() {
        indexed = null
    }

    private fun tag(f: RandomAccessFile, at: Long): String {
        val b = ByteArray(4)
        f.seek(at)
        f.readFully(b)
        return String(b, Charsets.ISO_8859_1)
    }

    private fun u16(f: RandomAccessFile, at: Long): Int {
        f.seek(at)
        return f.readUnsignedShort()
    }

    private fun u32(f: RandomAccessFile, at: Long): Long {
        f.seek(at)
        return f.readInt().toLong() and 0xffffffffL
    }
}
