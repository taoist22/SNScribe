package com.sndocx

import org.junit.Assert.assertEquals
import org.junit.Assert.assertFalse
import org.junit.Assert.assertTrue
import org.junit.Test

class SpellCheckerTest {
    private val en = SpellChecker.builtIn("en_US")

    @Test
    fun knowsWordsAndTheirForms() {
        val t0 = System.currentTimeMillis()
        SpellChecker.builtIn("en_US")
        println("loaded in ${System.currentTimeMillis() - t0} ms (cached)")
        for (w in listOf(
            "receive", "received", "receiving", "receives", "happiness", "unhappy", "rewrite", "running",
            "inflammation", "accounting", "debits", "credits", "psychology", "Psychology", "PSYCHOLOGY",
            "don't", "don’t", "Smith's", "children's", "studies", "quickly", "undo", "a", "I",
        )) assertTrue(w, en.check(w))
        for (w in listOf("recieve", "teh", "seperate", "definately", "occured", "accomodate", "psycology", "happyness")) {
            assertFalse(w, en.check(w))
        }
    }

    @Test
    fun suggestsTheLikelyWordFirst() {
        val cases = mapOf(
            "recieve" to "receive",
            "teh" to "the",
            "seperate" to "separate",
            "definately" to "definitely",
            "occured" to "occurred",
            "accomodate" to "accommodate",
            "Seperate" to "Separate",
        )
        for ((wrong, right) in cases) {
            val t0 = System.nanoTime()
            val s = en.suggest(wrong)
            val ms = (System.nanoTime() - t0) / 1_000_000
            println("$wrong → $s ($ms ms)")
            assertTrue("$wrong → $s", right in s.take(3))
        }
    }

    @Test
    fun twoWordsRunTogether() {
        assertTrue(en.suggest("alot").any { it == "a lot" })
        assertEquals(false, en.check("thecat"))
        assertTrue("the cat" in en.suggest("thecat"))
    }
}
