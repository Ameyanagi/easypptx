"""Tests for the fluent Deck builder (0.10.0)."""

import warnings

import pandas as pd
import pytest

from easypptx import Deck, Presentation

DF = pd.DataFrame({"Q": ["Q1", "Q2"], "Rev": [1, 2], "Cost": [3, 4]})


def slide_texts(slide):
    return [s.text_frame.text for s in slide.shapes if s.has_text_frame]


class TestChain:
    def test_full_chain_renders_expected_slides(self):
        pres = (
            Deck(theme="corporate")
            .title_slide("Q3 Review", subtitle="Finance", notes="opening")
            .slide("Highlights", notes="short")
            .bullets(["Revenue up 12%", ("EMEA strongest", 1)])
            .chart(DF, kind="column", value_columns=["Rev", "Cost"])
            .section("Details")
            .slide("Data")
            .table(DF)
            .text("Totals in appendix.")
            .build()
        )
        assert len(pres.slides) == 4
        assert pres.slides[0].notes == "opening"
        assert pres.slides[1].notes == "short"
        assert "Q3 Review" in slide_texts(pres.slides[0])
        assert "Details" in slide_texts(pres.slides[2])

    def test_every_method_returns_the_deck(self):
        deck = Deck()
        assert deck.slide("a") is deck
        assert deck.text("t") is deck
        assert deck.bullets(["b"]) is deck
        assert deck.chart(DF) is deck
        assert deck.table(DF) is deck
        assert deck.notes("n") is deck
        assert deck.tap(lambda s: None) is deck

    def test_save_writes_file(self, tmp_path):
        out = tmp_path / "deck.pptx"
        Deck().slide("One").text("hello").save(out)
        assert out.exists()

    def test_pres_deck_builds_onto_existing_presentation(self):
        pres = Presentation(theme="dark")
        result = pres.deck().slide("A").text("x").build()
        assert result is pres
        assert len(pres.slides) == 1


class TestValidation:
    def test_content_before_slide_raises(self):
        with pytest.raises(RuntimeError, match="Start a slide"):
            Deck().bullets(["a"])

    def test_content_after_title_slide_raises(self):
        with pytest.raises(RuntimeError, match="Start a slide"):
            Deck().title_slide("T").text("body")

    def test_unknown_chart_kind_raises_at_call_site(self):
        with pytest.raises(ValueError, match="starburst"):
            Deck().slide("x").chart(DF, kind="starburst")

    def test_bad_chart_data_raises_at_call_site(self):
        with pytest.raises(ValueError, match="Unsupported chart data type"):
            Deck().slide("x").chart(42)

    def test_missing_image_raises_at_call_site(self):
        with pytest.raises(FileNotFoundError):
            Deck().slide("x").image("nope.png")

    def test_bad_bullets_raise(self):
        with pytest.raises(TypeError):
            Deck().slide("x").bullets("not a list")
        with pytest.raises(TypeError):
            Deck().slide("x").bullets([("text", "level")])

    def test_consumed_deck_cannot_be_reused(self):
        deck = Deck().slide("a")
        deck.build()
        with pytest.raises(RuntimeError, match="already rendered"):
            deck.slide("b")
        with pytest.raises(RuntimeError, match="already rendered"):
            deck.build()


class TestLayout:
    def test_chart_expands_into_leftover_space(self):
        pres = Deck().slide("T").bullets(["one", "two"]).chart(DF).build()
        chart_frame = next(s for s in pres.slides[0].shapes if s.has_chart)
        # The chart should take a large share of the slide height
        assert chart_frame.height > pres.pptx_presentation.slide_height * 0.4

    def test_positioned_block_escapes_auto_layout(self):
        pres = Deck().slide("T").text("pinned", x=60, y=80, width=30, height=10).build()
        shape = next(s for s in pres.slides[0].shapes if "pinned" in s.text_frame.text)
        assert shape.top > pres.pptx_presentation.slide_height * 0.7

    def test_tap_receives_rendered_slide(self):
        seen = []
        Deck().slide("T").text("x").tap(lambda s: seen.append(len(s.shapes))).build()
        assert seen and seen[0] >= 2  # title + text already rendered

    def test_notes_append(self):
        pres = Deck().slide("T", notes="first").notes("second").build()
        assert pres.slides[0].notes == "first\nsecond"

    def test_rendering_is_warning_free(self):
        with warnings.catch_warnings():
            warnings.simplefilter("error")
            (
                Deck(theme="dark")
                .title_slide("T", subtitle="s")
                .slide("Content")
                .bullets(["a", "b", ("c", 1)])
                .chart(DF, value_columns=["Rev", "Cost"], show_values=True)
                .slide("More")
                .table(DF)
                .build()
            )


class TestPagination:
    def test_long_bullets_paginate_with_cont_titles(self):
        deck = Deck(theme="light").slide("Long list")
        deck.bullets([f"Bullet point number {i} with a reasonable amount of words" for i in range(40)])
        pres = deck.build()
        assert len(pres.slides) >= 2
        titles = [slide_texts(s)[0] for s in pres.slides]
        assert titles[0] == "Long list"
        assert all("(cont.)" in t for t in titles[1:])

    def test_notes_only_on_first_page(self):
        deck = Deck().slide("L", notes="only once")
        deck.bullets([f"item {i} with several words to wrap around" for i in range(40)])
        pres = deck.build()
        assert pres.slides[0].notes == "only once"
        assert pres.slides[1].notes == ""

    def test_paginate_false_compresses_instead(self):
        deck = Deck().slide("Dense", paginate=False)
        deck.bullets([f"item {i} with several words attached" for i in range(40)])
        pres = deck.build()
        assert len(pres.slides) == 1

    def test_mixed_blocks_flow_to_next_page(self):
        deck = Deck().slide("Mix")
        deck.bullets([f"filler line {i} with enough words to take space" for i in range(30)])
        deck.table(DF)
        pres = deck.build()
        assert len(pres.slides) >= 2
        # The table must exist on some page
        assert any(any(s.has_table for s in slide.shapes) for slide in pres.slides)
