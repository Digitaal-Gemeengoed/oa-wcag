package nl.overheid.wcag.check;

import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.junit.jupiter.api.Test;

import java.util.function.Consumer;

import static org.junit.jupiter.api.Assertions.assertEquals;

class UseOfColorCheckTest {

    static int meldingen(String gekleurd, Consumer<XWPFRun> extra) {
        XWPFDocument d = new XWPFDocument();
        XWPFParagraph p = d.createParagraph();
        p.createRun().setText("Mevrouw ");
        XWPFRun r = p.createRun();
        r.setText(gekleurd);
        r.setColor("1E6B33");
        extra.accept(r);
        p.createRun().setText(" meldde zich.");
        return new UseOfColorCheck().check(d, new DocumentContext(d)).size();
    }

    @Test
    void alleenKleurWordtGemeld() {
        assertEquals(1, meldingen("Jansen", r -> {}));
    }

    @Test
    void plaatshoudersVanDeAnonimiseringNiet() {
        assertEquals(0, meldingen("[PERSOON 11]", r -> {}));
        assertEquals(0, meldingen("[E-MAIL 1]", r -> {}));
        assertEquals(0, meldingen("[IBAN 1]", r -> {}));
        assertEquals(1, meldingen("[zie bijlage]", r -> {}), "gewone tekst tussen haken telt niet als plaatshouder");
    }

    @Test
    void elkAnderKenmerkTelt() {
        assertEquals(0, meldingen("Jansen", r -> r.setUnderline(UnderlinePatterns.DOTTED)));
        assertEquals(0, meldingen("Jansen", r -> r.setStrikeThrough(true)));
        assertEquals(0, meldingen("Jansen", r -> r.setBold(true)));
    }
}
