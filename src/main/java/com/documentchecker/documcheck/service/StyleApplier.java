package com.documentchecker.documcheck.service; 

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import java.math.BigInteger;

public class StyleApplier {

    public static void applyGostStyles(XWPFDocument document, int fontSize) {
        if (document == null) return;

        for (XWPFParagraph paragraph : document.getParagraphs()) {
            applyParagraphStyles(paragraph, fontSize);
        }
        setGostMargins(document);
    }

    private static void applyParagraphStyles(XWPFParagraph paragraph, int fontSize) {
        paragraph.setAlignment(ParagraphAlignment.BOTH);
        
        setLineSpacing(paragraph, 360L); 

        paragraph.setIndentationFirstLine(709);

        
        for (XWPFRun run : paragraph.getRuns()) {
            run.setFontFamily("Times New Roman");
            run.setFontSize(fontSize);
            run.setColor("000000"); 
            run.setItalic(false);
            run.setBold(false);
        }
    }

    private static void setLineSpacing(XWPFParagraph paragraph, long spacingValue) {
        CTPPr ppr = paragraph.getCTP().isSetPPr() ? paragraph.getCTP().getPPr() : paragraph.getCTP().addNewPPr();
        CTSpacing spacing = ppr.isSetSpacing() ? ppr.getSpacing() : ppr.addNewSpacing();
        spacing.setLine(BigInteger.valueOf(spacingValue));
        spacing.setLineRule(STLineSpacingRule.AUTO); 
    }

    private static void setGostMargins(XWPFDocument document) {
        CTSectPr sectPr = document.getDocument().getBody().isSetSectPr() ? 
                          document.getDocument().getBody().getSectPr() : 
                          document.getDocument().getBody().addNewSectPr();
                          
        CTPageMar pageMar = sectPr.isSetPgMar() ? sectPr.getPgMar() : sectPr.addNewPgMar();

        pageMar.setLeft(BigInteger.valueOf(1701));  
        pageMar.setRight(BigInteger.valueOf(850));   
        pageMar.setTop(BigInteger.valueOf(1134));    
        pageMar.setBottom(BigInteger.valueOf(1134)); 
    }
}