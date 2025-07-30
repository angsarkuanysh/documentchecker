package com.documentchecker.documcheck.service; 

import java.io.InputStream;
import java.math.BigInteger;
import java.util.HashSet;
import java.util.List;
import java.util.Locale;
import java.util.Set;

import org.apache.poi.xwpf.usermodel.*;
import org.springframework.web.multipart.MultipartFile;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;


public class DocxToHtmlConverter {

    private static final double PAGE_HEIGHT_PT = 730.0;
    private static final double AVG_CHAR_HEIGHT_PT = 14.0;
    private static final int AVG_CHARS_PER_LINE = 85;
    
    double currentPageHeight = 0;
    
    public String convertDocxToHtmlWithErrors(MultipartFile file, double newIndent, double newLineSpacing, int newFontSize) throws Exception {
        try (InputStream is = file.getInputStream(); XWPFDocument doc = new XWPFDocument(is)) {
            StringBuilder html = new StringBuilder();
            String headerHtml = extractHeaders(doc);
            String footerHtml = extractFooters(doc);
            html.append(headerHtml);

            List<IBodyElement> elements = doc.getBodyElements();
            
            BigInteger currentListNumId = null;
            String currentListTag = null;

            for (IBodyElement el : elements) {
                Set<String> paragraphErrors = new HashSet<>();
                if (el instanceof XWPFParagraph) {
                    XWPFParagraph p = (XWPFParagraph) el;
                    BigInteger numId = p.getNumID();

                    if (numId != null) {
                        if (!numId.equals(currentListNumId)) {
                            if (currentListNumId != null) {
                                html.append("</").append(currentListTag).append(">\n");
                            }
                            currentListNumId = numId;
                            currentListTag = getListTag(p); 
                            html.append("<").append(currentListTag).append(">\n");
                        }
                        html.append(processListItem(p, paragraphErrors, newFontSize));

                    } else {
                        if (currentListNumId != null) {
                            html.append("</").append(currentListTag).append(">\n");
                            currentListNumId = null; 
                            currentListTag = null;
                        }

                        String style = p.getStyle();
                        if (style != null && style.matches("Heading[1-6]")) {
                            html.append(processHeading(p, paragraphErrors, newFontSize));
                        } else {
                            html.append(processParagraph(p, headerHtml, footerHtml, paragraphErrors, false, newFontSize, newIndent, newLineSpacing));
                        }
                    }
                } else if (el instanceof XWPFTable) {
                    if (currentListNumId != null) {
                        html.append("</").append(currentListTag).append(">\n");
                        currentListNumId = null;
                        currentListTag = null;
                    }
                    html.append(processTable((XWPFTable) el, headerHtml, footerHtml, paragraphErrors, newFontSize, newIndent, newLineSpacing));
                }
            }

            if (currentListNumId != null) {
                html.append("</").append(currentListTag).append(">\n");
            }
            
            
            return html.toString();
        }
    }

    private String processHeading(XWPFParagraph p, Set<String> paragraphErrors, int fontSize) {
        int level = Integer.parseInt(p.getStyle().substring("Heading".length()));
        String tag = "h" + level;
        String align = getAlign(p.getAlignment());
        StringBuilder sb = new StringBuilder();
        sb.append("<").append(tag).append(" style='text-align:")
          .append(align).append(";'>");
        for (XWPFRun r : p.getRuns()) {
            sb.append(processRun(r, paragraphErrors, fontSize));
        }
        sb.append("</").append(tag).append(">\n");
        return sb.toString();
    }

    private String processPicturesFromParagraph(XWPFParagraph p) throws Exception {
        StringBuilder sb = new StringBuilder();

        for (XWPFRun run : p.getRuns()) {
            List<XWPFPicture> pictures = run.getEmbeddedPictures();
            for (XWPFPicture picture : pictures) {
                XWPFPictureData pictureData = picture.getPictureData();
                if (pictureData == null) continue;

                String base64 = java.util.Base64.getEncoder().encodeToString(pictureData.getData());
                String extension = pictureData.suggestFileExtension().toLowerCase(Locale.ROOT);
                String mimeType = getImageMimeType(extension);

                sb.append("<img src='data:").append(mimeType).append(";base64,")
                .append(base64).append("' style='max-width:100%; height:auto;' />\n");
            }
        }

        return sb.toString();
    }

    private String getImageMimeType(String extension) {
        return switch (extension) {
            case "emf" -> "image/x-emf";
            case "wmf" -> "image/x-wmf";
            case "pict" -> "image/x-pict";
            case "jpeg", "jpg" -> "image/jpeg";
            case "png" -> "image/png";
            case "dib", "bmp" -> "image/bmp";
            case "gif" -> "image/gif";
            case "tiff", "tif" -> "image/tiff";
            default -> "application/octet-stream";
        };
    }
    
    private String getListTag(XWPFParagraph p) {
        try {
            BigInteger numId = p.getNumID();
            if (numId == null) return "ul"; 

            XWPFDocument doc = p.getDocument();
            XWPFNumbering numbering = doc.getNumbering();
            if (numbering == null) return "ul";

            XWPFNum num = numbering.getNum(numId);
            if (num == null) return "ul";

            BigInteger abstractNumId = num.getCTNum().getAbstractNumId().getVal();
            if (abstractNumId == null) return "ul";

            XWPFAbstractNum absNum = numbering.getAbstractNum(abstractNumId);
            if (absNum == null || absNum.getCTAbstractNum() == null) return "ul";
            
            CTAbstractNum ctAbsNum = absNum.getCTAbstractNum();

            BigInteger ilvl = p.getNumIlvl();
            int level = (ilvl != null) ? ilvl.intValue() : 0;
            
            if (ctAbsNum.sizeOfLvlArray() <= level) {
                return "ul";
            }

            CTLvl ctLvl = ctAbsNum.getLvlArray(level);
            if (ctLvl == null || ctLvl.getNumFmt() == null || ctLvl.getNumFmt().getVal() == null) {
                return "ul";
            }

            STNumberFormat.Enum fmt = ctLvl.getNumFmt().getVal();
            
            return "decimal".equals(fmt.toString()) ? "ol" : "ul";

        } catch (Exception e) {
            e.printStackTrace(); 
            return "ul"; 
        }
    }
    
    private String processListItem(XWPFParagraph p, Set<String> paragraphErrors, int fontSize) {
        StringBuilder sb = new StringBuilder();
        
        int level = p.getNumIlvl() != null ? p.getNumIlvl().intValue() : 0;
        
        sb.append("<li style='margin-left: ").append(level * 1.5).append("em;'>");
        
        

        for (XWPFRun r : p.getRuns()) {
            sb.append(processRun(r, paragraphErrors, fontSize));
        }
        
        sb.append("</li>\n");
        return sb.toString();
    }

    private String  processRun(XWPFRun r, Set<String> paragraphErrors, int fontSize) {
        StringBuilder sb = new StringBuilder();
        String text = r.text();
        System.out.println(text);
        if (text == null || text.isEmpty()) return "";
        if (r.getCTR() != null && r.getCTR().getRPr() != null && 
            r.getCTR().getRPr().getSzList() != null && !r.getCTR().getRPr().getSzList().isEmpty()) { 
                int fontSizePts = r.getFontSize();
                if (fontSizePts != -1 && fontSizePts != fontSize) {
                    
            
                    paragraphErrors.add("Неверный размер шрифта: " + fontSizePts + "pt (ожидается " + fontSize + "pt)");
                }
        } 
        
        String fontFamily = r.getFontFamily();

        if (fontFamily != null && !"Times New Roman".equalsIgnoreCase(fontFamily)) {
            paragraphErrors.add("Неверный шрифт: " + fontFamily);
        }

        fontFamily = r.getFontFamily() != null ? r.getFontFamily() : "Times New Roman";
        String font = r.getFontFamily() == null ? "Times New Roman" : r.getFontFamily();
        
        int size = r.getFontSize() <= 0 ? 14 : r.getFontSize();
        
        sb.append("<span style='font-family:").append(font).append("; font-size:")
          .append(size).append("pt;");
        
          if (!paragraphErrors.isEmpty()) {
                        sb.append(" background-color: yellow; border-bottom: 1px dashed red; cursor: help;'")
                            .append(" title='").append(String.valueOf(paragraphErrors)).append("'");
                    } else {
                        sb.append("'");
                    }
        sb.append(">");

        if (r.isBold()) sb.append("<b>");
        if (r.isItalic()) sb.append("<i>");
        if (r.getUnderline() != UnderlinePatterns.NONE) sb.append("<u>");

        sb.append(escapeHtml(text));

        if (r.getUnderline() != UnderlinePatterns.NONE) sb.append("</u>");
        if (r.isItalic()) sb.append("</i>");
        if (r.isBold()) sb.append("</b>");
        sb.append("</span>");

        return sb.toString();
    }

    private String processParagraph(XWPFParagraph p, String headerHtml, String footerHtml, Set<String> paragraphErrors, boolean isInsideTable, int newFontSize, double newIndent, double newLineSpacing) {
        boolean isHeading = isHeading(p);
        boolean isTitlePageElement = isTitlePageElement(p);
        boolean isListItem = isListItem(p);
        
        if (!isInsideTable) {
            if (isTitlePageElement || isListItem) {
            } else if (isHeading) {
                
                    if (p.getAlignment() != ParagraphAlignment.CENTER) {
                        paragraphErrors.add("Заголовок не выровнен по центру");
                    }
                
            } else {
                
                if (p.getAlignment() != ParagraphAlignment.BOTH) {
                    paragraphErrors.add("Текст не выровнен по ширине");
                }
                int indent = getParagraphFirstLineIndent(p);
                double cm = newIndent * 567.0;
                if (!isListItem(p) && (indent < cm - 40 || indent > cm + 40)) {
                    paragraphErrors.add("Неверный отступ первой строки: " + indent + " (ожидается ~" + newIndent + "см)");
                }
                if (p.getSpacingBetween() != -1 || p.getSpacingBetween() != newLineSpacing) {
                    paragraphErrors.add("Неверный межстрочный интервал: " + p.getSpacingBetween() + " (ожидается "+ newLineSpacing +")");
                }
            }
        }

        String align = getAlign(p.getAlignment());
        String indent = isInsideTable ? "0" : twipToCm(p.getFirstLineIndent());

        StringBuilder sb = new StringBuilder();

        if (!isInsideTable) {
            double paragraphHeight = estimateParagraphHeight(p);
            if (currentPageHeight + paragraphHeight > PAGE_HEIGHT_PT && currentPageHeight > 0) {
                sb.append(footerHtml);
                sb.append("<div class='page-break'></div>");
                currentPageHeight = 0;
                sb.append(headerHtml);
            }
            currentPageHeight += paragraphHeight;
        }

        if (p.getText().trim().isEmpty() && p.getRuns().isEmpty()) {
            if (!isInsideTable) {
                sb.append("<p style='margin:0; padding:0; height:").append(AVG_CHAR_HEIGHT_PT).append("pt;'>&nbsp;</p>");
                currentPageHeight += AVG_CHAR_HEIGHT_PT;
            } else {
                sb.append("<p style='margin:0; padding:0;'>&nbsp;</p>");
            }
            return sb.toString();
        }

        sb.append("<p style='text-align:")
        .append(align).append("; text-indent:")
        .append(indent).append("cm; margin: 0; padding: 0;'>");

        try {
            sb.append(processPicturesFromParagraph(p));
        } catch (Exception e) {
            e.printStackTrace(); 
        }
        for (XWPFRun r : p.getRuns()) {
            sb.append(processRun(r, paragraphErrors, newFontSize));
        }

        sb.append("</p>\n");
        return sb.toString();
    }

    private String twipToCm(int tw) {
        double cm = tw / 567.0;
        return String.format(Locale.US, "%.2f", cm);
    }

    private String processTable(XWPFTable table, String headerHtml, String footerHtml, Set<String> paragraphErrors, int newFontSize, double newIndent, double newLineSpacing) {
        
        StringBuilder html = new StringBuilder();
        html.append("<table border='1' style='border-collapse: collapse; width: 100%;'>");

        for (XWPFTableRow row : table.getRows()) {
            html.append("<tr>");
            for (XWPFTableCell cell : row.getTableCells()) {

                html.append("<td style='padding: 5px;'>");
                for (XWPFParagraph paragraph : cell.getParagraphs()) {
                    html.append(processParagraph(paragraph, headerHtml, footerHtml, paragraphErrors, true, newFontSize, newIndent, newLineSpacing));
                }

                html.append("</td>");
            }
            html.append("</tr>");
        }

        html.append("</table>\n");
        return html.toString();
    }

    private double estimateParagraphHeight(XWPFParagraph para) {
        double height = 0;
        height += convertTwipsToPt(para.getSpacingBefore());
        height += convertTwipsToPt(para.getSpacingAfter());

        String text = para.getText();
        if (!text.isEmpty()) {
            double lineSpacing = para.getSpacingBetween() > 0 ? para.getSpacingBetween() : 1.0;
            int lineCount = (int) Math.ceil((double) text.length() / AVG_CHARS_PER_LINE);
            if (lineCount == 0) lineCount = 1;
            height += lineCount * AVG_CHAR_HEIGHT_PT * lineSpacing;
        } else {
            height += AVG_CHAR_HEIGHT_PT;
        }
        return height;
    }
    
    private String extractHeaders(XWPFDocument doc) {
        StringBuilder headersHtml = new StringBuilder();
        if (doc.getHeaderList() != null && !doc.getHeaderList().isEmpty()) {
            XWPFHeader header = doc.getHeaderList().get(0);
            headersHtml.append("<div class='header'>");
            for (XWPFParagraph p : header.getParagraphs()) {
                headersHtml.append("<p>").append(escapeHtml(p.getText())).append("</p>");
            }
            headersHtml.append("</div>");
        }
        return headersHtml.toString();
    }

    private String extractFooters(XWPFDocument doc) {
        StringBuilder footersHtml = new StringBuilder();
        if (doc.getFooterList() != null && !doc.getFooterList().isEmpty()) {
            XWPFFooter footer = doc.getFooterList().get(0);
            footersHtml.append("<div class='footer'>");
            for (XWPFParagraph p : footer.getParagraphs()) {
                footersHtml.append("<p>").append(escapeHtml(p.getText())).append("</p>");
            }
            footersHtml.append("</div>");
        }
        return footersHtml.toString();
    }

    private String escapeHtml(String text) {
        return text.replace("&", "&amp;")
                   .replace("<", "&lt;")
                   .replace(">", "&gt;");
    }

    private String getAlign(ParagraphAlignment align) {
        if (align == ParagraphAlignment.CENTER) return "center";
        if (align == ParagraphAlignment.RIGHT) return "right";
        if (align == ParagraphAlignment.BOTH) return "justify";
        return "left";
    }

    private double convertTwipsToPt(int twips) {
        if (twips <= 0) return 0;
        return twips / 20.0;
    }
    
    private boolean isListItem(XWPFParagraph para) {
        String text = para.getText().trim();
        return text.startsWith("-") || text.startsWith("*") || (text.matches(".*\\.{3,}.*") && text.matches(".*\\d$"));
    }

    private boolean isTitlePageElement(XWPFParagraph para) {
        return para.getDocument().getPosOfParagraph(para) < 20; 
    }

    private boolean isHeading(XWPFParagraph para) {
        String text = para.getText().trim();
        if (text.isEmpty() || text.length() > 40 || isListItem(para) || isTitlePageElement(para)) {
            return false;
        }
        boolean isAllBold = para.getRuns().stream().allMatch(XWPFRun::isBold);
        if (isAllBold && !text.endsWith(".") && para.getAlignment() == ParagraphAlignment.CENTER) {
            return true;
        }

        if (text.equals(text.toUpperCase()) || text.length() > 5 || !text.matches(".*\\d.*")) {
             return true;
        }
        
        return false;
    }

    private int getParagraphFirstLineIndent(XWPFParagraph para) {
        if (para.getFirstLineIndent() != -1) {
            return para.getFirstLineIndent();
        }

        if (para.getStyleID() != null) {
            XWPFStyle style = para.getDocument().getStyles().getStyle(para.getStyleID());

            if (style != null && style.getCTStyle() != null && style.getCTStyle().getPPr() != null) {
                if (style.getCTStyle().getPPr() instanceof CTPPr) {
                        CTPPr ppr = (CTPPr) style.getCTStyle().getPPr();
                    
                

                if (ppr.getInd() != null) {
                    CTInd ind = ppr.getInd();

                    if (ind.isSetFirstLine()) {
                        Object firstLineRaw = ind.getFirstLine();

                        if (firstLineRaw instanceof BigInteger) {
                            return ((BigInteger) firstLineRaw).intValue();
                        }
                    }
                }
            }
            }
        }
        return -1; 
    }
}
