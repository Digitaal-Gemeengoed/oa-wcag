package nl.overheid.wcag.check;

import nl.overheid.wcag.model.AccessibilityFinding;
import nl.overheid.wcag.model.ElementLocation;
import org.apache.poi.xwpf.usermodel.*;

import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder;

import java.util.ArrayList;
import java.util.regex.Pattern;
import java.util.List;

/**
 * WCAG 1.4.1 (Niveau A): Gebruik van kleur.
 * Detecteert gekleurde tekst zonder extra opmaak (vet/cursief)
 * die erop kan wijzen dat kleur als enige onderscheider wordt gebruikt.
 * Telt als extra kenmerk: vet, cursief, elke onderstreping, doorhaling, een
 * kader om de tekst, en de plaatshouders van de anonimisering ([PERSOON 1]).
 */
public class UseOfColorCheck implements AccessibilityCheck {

    @Override
    public String getCheckId() { return "USE_OF_COLOR_CHECK"; }

    @Override
    public String getDisplayName() { return "Gebruik van kleur"; }

    @Override
    public List<AccessibilityFinding> check(XWPFDocument document, DocumentContext context) {
        List<AccessibilityFinding> findings = new ArrayList<>();

        for (XWPFParagraph paragraph : document.getParagraphs()) {
            List<XWPFRun> runs = paragraph.getRuns();
            if (runs.size() < 2) continue; // Minimaal 2 runs nodig voor kleurverschil

            boolean hasColoredRunWithoutFormatting = false;
            String coloredText = null;

            for (XWPFRun run : runs) {
                String text = run.getText(0);
                if (text == null || text.trim().isEmpty()) continue;
                if (!isColored(run) || hasOtherIndicator(run, text)) continue;
                hasColoredRunWithoutFormatting = true;
                coloredText = text;
                break;
            }

            if (hasColoredRunWithoutFormatting) {
                int paraIdx = context.getParagraphIndex(paragraph);
                int bodyIdx = context.getParagraphBodyElementIndex(paragraph);
                findings.add(AccessibilityFinding.builder(getCheckId())
                        .warning()
                        .wcag("1.4.1", "A")
                        .message("Gekleurde tekst zonder extra visuele aanduiding (vet/cursief/onderstreept). "
                                + "Kleur wordt mogelijk als enige onderscheider gebruikt.")
                        .suggestion("Voeg naast kleur ook een ander visueel kenmerk toe "
                                + "(bijv. vet, cursief of onderstreept) zodat de informatie "
                                + "ook zonder kleur herkenbaar is.")
                        .location(ElementLocation.paragraph(bodyIdx, paraIdx, coloredText))
                        .build());
            }
        }

        return findings;
    }

    /**
     * Plaatshouder van de anonimisering, zoals [PERSOON 11] of [E-MAIL 1]: de
     * tekst zelf zegt wat het is; de kleur (het PPL-niveau) is een extra.
     */
    static final Pattern PLAATSHOUDER = Pattern.compile("\\s*\\[[\\p{Lu}][\\p{Lu}\\d\\- ]* \\d+\\]\\s*");

    private static boolean isColored(XWPFRun run) {
        String color = run.getColor();
        return color != null && !color.isEmpty()
                && !"000000".equalsIgnoreCase(color)
                && !"auto".equalsIgnoreCase(color);
    }

    /** Vet, cursief, elke onderstreping, doorhaling, een kader om de tekst, of een plaatshouder. */
    private static boolean hasOtherIndicator(XWPFRun run, String text) {
        if (run.isBold() || run.isItalic() || run.isStrikeThrough() || run.isDoubleStrikeThrough()) return true;
        UnderlinePatterns u = run.getUnderline();
        if (u != null && u != UnderlinePatterns.NONE) return true;
        CTRPr rpr = run.getCTR().getRPr();
        if (rpr != null && rpr.sizeOfBdrArray() > 0) {
            STBorder.Enum soort = rpr.getBdrArray(0).getVal();
            if (soort != null && soort != STBorder.NONE && soort != STBorder.NIL) return true;
        }
        return PLAATSHOUDER.matcher(text).matches();
    }
}
