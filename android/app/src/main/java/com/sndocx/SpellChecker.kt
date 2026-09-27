package com.sndocx

import java.util.regex.Pattern

/**
 * A spell checker for Hunspell dictionaries (the .aff/.dic pairs LibreOffice and Firefox use),
 * covering what word lists like SCOWL's en_US need: prefix and suffix rules with their
 * conditions, cross products (a prefix and a suffix together), NEEDAFFIX, FORBIDDENWORD,
 * ONLYINCOMPOUND, NOSUGGEST, ICONV, and REP/TRY for suggestions. Compounding (German,
 * Dutch …) is not covered: such words may be flagged.
 */
class SpellChecker(aff: String, dic: String) {
    private class Rule(val strip: String, val add: String, val cond: Pattern?)
    private class Group(val cross: Boolean, val rules: MutableList<Rule> = ArrayList())

    private val words = HashMap<String, String>()
    private val prefixes = LinkedHashMap<Char, Group>()
    private val suffixes = LinkedHashMap<Char, Group>()
    private val rep = ArrayList<Pair<String, String>>()
    private val iconv = ArrayList<Pair<String, String>>()
    private var tryChars = "esianrtolcdugmphbyfvkwz'"
    private var needAffix: Char? = null
    private var forbidden: Char? = null
    private var onlyInCompound: Char? = null
    private var noSuggest: Char? = null

    init {
        for (raw in aff.lineSequence()) {
            val line = raw.trim()
            if (line.isEmpty() || line.startsWith("#")) continue
            val f = line.split(Regex("\\s+"))
            when (f[0]) {
                "TRY" -> tryChars = f.getOrElse(1) { tryChars }
                "NEEDAFFIX" -> needAffix = f.getOrNull(1)?.firstOrNull()
                "FORBIDDENWORD" -> forbidden = f.getOrNull(1)?.firstOrNull()
                "ONLYINCOMPOUND" -> onlyInCompound = f.getOrNull(1)?.firstOrNull()
                "NOSUGGEST" -> noSuggest = f.getOrNull(1)?.firstOrNull()
                "REP" -> if (f.size >= 3) rep.add(f[1].replace('_', ' ') to f[2].replace('_', ' '))
                "ICONV" -> if (f.size >= 3) iconv.add(f[1] to f[2])
                "PFX", "SFX" -> {
                    val map = if (f[0] == "PFX") prefixes else suffixes
                    val flag = f[1].first()
                    if (f.size == 4 && (f[2] == "Y" || f[2] == "N") && f[3].all { it.isDigit() }) {
                        map[flag] = Group(f[2] == "Y")
                    } else if (f.size >= 4) {
                        val group = map.getOrPut(flag) { Group(false) }
                        val strip = if (f[2] == "0") "" else f[2]
                        val add = f[3].substringBefore('/').let { if (it == "0") "" else it }
                        val cond = f.getOrNull(4)?.takeIf { it != "." }
                        group.rules.add(
                            Rule(strip, add, cond?.let { Pattern.compile(if (f[0] == "PFX") "^$it.*" else ".*$it$") }),
                        )
                    }
                }
            }
        }
        var first = true
        for (raw in dic.lineSequence()) {
            val line = raw.trim()
            if (line.isEmpty()) continue
            if (first) {
                first = false
                if (line.all { it.isDigit() }) continue
            }
            // "word/FLAGS" (a "\/" is a literal slash); anything after whitespace is morphology.
            val entry = line.split(Regex("[\\t ]"), limit = 2)[0]
            val slash = entry.indexOf('/')
            val word = if (slash >= 0) entry.substring(0, slash) else entry
            val flags = if (slash >= 0) entry.substring(slash + 1) else ""
            words[word] = (words[word] ?: "") + flags
        }
    }

    private fun conv(w: String): String = iconv.fold(w) { s, (a, b) -> s.replace(a, b) }

    private fun usable(flags: String?): Boolean =
        flags != null && listOfNotNull(needAffix, forbidden, onlyInCompound).none { it in flags }

    private fun stemHas(stem: String, flag: Char): Boolean {
        val f = words[stem] ?: return false
        return flag in f && forbidden?.let { it !in f } != false
    }

    /** The word as written, taking suffixes and prefixes (and both together) off to find its stem. */
    private fun lookup(w: String): Boolean {
        val direct = words[w]
        if (direct != null) {
            if (forbidden != null && forbidden!! in direct) return false
            if (usable(direct)) return true
        }
        for ((flag, g) in suffixes) {
            for (r in g.rules) {
                if (!w.endsWith(r.add) || w.length <= r.add.length) continue
                val base = w.substring(0, w.length - r.add.length) + r.strip
                if (r.cond != null && !r.cond.matcher(base).matches()) continue
                if (stemHas(base, flag)) return true
                if (g.cross && prefixed(base) { stem, pflag -> stemHas(stem, pflag) && stemHas(stem, flag) }) return true
            }
        }
        return prefixed(w) { stem, pflag -> stemHas(stem, pflag) }
    }

    private inline fun prefixed(w: String, ok: (String, Char) -> Boolean): Boolean {
        for ((flag, g) in prefixes) {
            for (r in g.rules) {
                if (!w.startsWith(r.add) || w.length <= r.add.length) continue
                val base = r.strip + w.substring(r.add.length)
                if (r.cond != null && !r.cond.matcher(base).matches()) continue
                if (ok(base, flag)) return true
            }
        }
        return false
    }

    /** Whether [word] is spelled right: as written, or lower-cased when capitalised or all in capitals. */
    fun check(word: String): Boolean {
        val w = conv(word.trim())
        if (w.isEmpty()) return true
        if (lookup(w)) return true
        val lower = w.lowercase()
        if (w != lower && (w.drop(1) == w.drop(1).lowercase() || w == w.uppercase())) {
            if (lookup(lower)) return true
            if (w == w.uppercase() && lookup(lower.replaceFirstChar { it.uppercase() })) return true
        }
        return false
    }

    private fun suggestible(w: String): Boolean {
        // "a lot" (a REP pattern): each word must be right.
        if (' ' in w) return w.split(' ').all { it.isNotEmpty() && suggestible(it) }
        if (!check(w)) return false
        val f = words[w] ?: words[w.lowercase()]
        return noSuggest == null || f == null || noSuggest!! !in f
    }

    /**
     * Up to [max] corrections: the dictionary's REP patterns first (common misspellings),
     * then one-letter edits (a letter swapped, dropped, added or changed), two words run
     * together, then two-letter edits when nothing closer is found. Capitalisation follows
     * the word.
     */
    fun suggest(word: String, max: Int = 6): List<String> {
        val w = conv(word.trim())
        if (w.isEmpty()) return emptyList()
        val capital = w.first().isUpperCase() && w.drop(1) == w.drop(1).lowercase()
        val caps = w.length > 1 && w == w.uppercase()
        val base = if (capital || caps) w.lowercase() else w
        val out = LinkedHashSet<String>()
        fun consider(c: String) {
            if (c != base && c.isNotEmpty() && out.size < max * 3 && suggestible(c)) out.add(c)
        }
        for ((from, to) in rep) {
            var i = base.indexOf(from)
            while (i >= 0) {
                consider(base.substring(0, i) + to + base.substring(i + from.length))
                i = base.indexOf(from, i + 1)
            }
        }
        val edits1 = edits(base)
        edits1.forEach(::consider)
        for (i in 1 until base.length) {
            val a = base.substring(0, i)
            val b = base.substring(i)
            if (a.length > 1 && b.length > 1 && check(a) && check(b)) out.add("$a $b")
        }
        if (out.isEmpty()) {
            var budget = 60000
            loop@ for (e in edits1) for (e2 in edits(e)) {
                if (--budget < 0) break@loop
                consider(e2)
                if (out.size >= max) break@loop
            }
        }
        // Closest first: same first letter and length nearest the word.
        val ranked = out.sortedWith(compareBy<String>({ if (it.firstOrNull() == base.firstOrNull()) 0 else 1 }, { kotlin.math.abs(it.length - base.length) }))
        return ranked.take(max).map { s ->
            when {
                caps -> s.uppercase()
                capital -> s.replaceFirstChar { it.uppercase() }
                else -> s
            }
        }
    }

    private fun edits(w: String): List<String> {
        val out = ArrayList<String>()
        for (i in w.indices) out.add(w.removeRange(i, i + 1))
        for (i in 0 until w.length - 1) out.add(w.substring(0, i) + w[i + 1] + w[i] + w.substring(i + 2))
        for (i in w.indices) for (c in tryChars) if (c != w[i]) out.add(w.substring(0, i) + c + w.substring(i + 1))
        for (i in 0..w.length) for (c in tryChars) out.add(w.substring(0, i) + c + w.substring(i))
        return out
    }

    companion object {
        private val cache = HashMap<String, SpellChecker>()

        /** The built-in dictionary for [lang] (only "en_US" so far), loaded once. */
        @Synchronized
        fun builtIn(lang: String): SpellChecker {
            cache[lang]?.let { return it }
            val loader = SpellChecker::class.java.classLoader ?: error("no class loader")
            fun read(name: String) = loader.getResourceAsStream("dict/$name")?.use { it.readBytes().toString(Charsets.UTF_8) }
                ?: error("no built-in dictionary $name")
            return SpellChecker(read("$lang.aff"), read("$lang.dic")).also { cache[lang] = it }
        }
    }
}
