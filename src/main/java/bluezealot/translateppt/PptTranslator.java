package bluezealot.translateppt;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFDiagram;
import org.apache.poi.xslf.usermodel.XSLFGroupShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

import lombok.extern.slf4j.Slf4j;

@Slf4j
public class PptTranslator {
    private final Translator translator;
    
    public PptTranslator(Translator translator) {
        this.translator = translator;
    }

    public void translatePpt(String sourcePptPath, String targetPptPath, String sourceLang, String targetLang) throws IOException {
        try(var fis = new FileInputStream(sourcePptPath)){
            XMLSlideShow ppt = new XMLSlideShow(fis);
            
            // Collect all text shapes first, including those in group shapes
            List<XSLFTextShape> textShapes = new ArrayList<>();
            
            // Process shapes recursively
            for (XSLFSlide slide : ppt.getSlides()) {
                List<XSLFShape> shapesToProcess = new ArrayList<>();
                shapesToProcess.addAll(slide.getShapes());
                
                while (!shapesToProcess.isEmpty()) {
                    XSLFShape shape = shapesToProcess.remove(0);
                    if (shape instanceof XSLFTextShape) {
                        textShapes.add((XSLFTextShape) shape);
                    } else if (shape instanceof XSLFGroupShape) {
                        shapesToProcess.addAll(0, ((XSLFGroupShape) shape).getShapes());
                    } else if (shape instanceof XSLFTable) {
                        // Process table cells
                        XSLFTable table = (XSLFTable) shape;
                        for (int row = 0; row < table.getNumberOfRows(); row++) {
                            for (int col = 0; col < table.getNumberOfColumns(); col++) {
                                XSLFTextShape cellShape = table.getCell(row, col);
                                if (cellShape != null) {
                                    textShapes.add(cellShape);
                                }
                            }
                        }
                    } else if (shape instanceof XSLFDiagram) {
                        // Process diagram shapes
                        XSLFDiagram diagram = (XSLFDiagram) shape;
                        XSLFDiagram.XSLFDiagramGroupShape diagramGroupShape = diagram.getGroupShape();
                        if (diagramGroupShape != null) {
                            List<XSLFShape> diagramShapes = diagramGroupShape.getShapes();
                            for (XSLFShape diagramShape : diagramShapes) {
                                if (diagramShape instanceof XSLFTextShape) {
                                    // Store the diagram shape as well as the text shape
                                    textShapes.add((XSLFTextShape) diagramShape);
                                }
                            }
                        }
                    } else if (shape instanceof XSLFGroupShape) {
                        // Process group shapes
                        XSLFGroupShape group = (XSLFGroupShape) shape;
                        List<XSLFShape> groupShapes = group.getShapes();
                        if (groupShapes != null) {
                            for (XSLFShape groupShape : groupShapes) {
                                if (groupShape instanceof XSLFTextShape) {
                                    textShapes.add((XSLFTextShape) groupShape);
                                } else if (groupShape instanceof XSLFGroupShape) {
                                    shapesToProcess.add(0, groupShape);
                                }
                            }
                        }
                    }
                }
            }
            
            // Process collected text shapes
            for (XSLFTextShape shape : textShapes) {
                // Process each paragraph separately
                for (XSLFTextParagraph paragraph : shape.getTextParagraphs()) {
                    StringBuilder paragraphText = new StringBuilder();
                    List<XSLFTextRun> runsInParagraph = new ArrayList<>();
                    
                    // Collect text and runs for this paragraph
                    for (XSLFTextRun run : paragraph.getTextRuns()) {
                        String runText = run.getRawText();
                        if (runText != null && !runText.trim().isEmpty()) {
                            paragraphText.append(runText);
                            runsInParagraph.add(run);
                        }
                    }
                    
                    String originalParagraphText = paragraphText.toString();
                    if (!originalParagraphText.trim().isEmpty()) {
                        log.info("Translating paragraph: {}", originalParagraphText);
                        try {
                            // Translate the paragraph
                            String translatedParagraph = translator.translate(originalParagraphText, sourceLang, targetLang);
                            
                            // Split translated text by spaces to match runs
                            String[] translatedParts = translatedParagraph.split("\\s+");
                            int partIndex = 0;
                            
                            // Distribute translated text back to runs in this paragraph
                            for (XSLFTextRun run : runsInParagraph) {
                                if (partIndex < translatedParts.length) {
                                    String translatedPart = translatedParts[partIndex++];
                                    run.setText(translatedPart);
                                } else {
                                    // If we run out of translated parts, just clear the remaining runs
                                    run.setText("");
                                }
                            }
                            
                            log.info("Translated to: {}", translatedParagraph);
                        } catch (Exception e) {
                            log.error("Error translating paragraph:", e);
                            // Restore original text for runs in this paragraph if translation fails
                            for (XSLFTextRun run : runsInParagraph) {
                                run.setText(run.getRawText());
                            }
                        }
                    }
                }
            }
            
            // Save the translated PPT
            try(FileOutputStream fos = new FileOutputStream(targetPptPath, false)) {
                ppt.write(fos);
                log.info("Successfully saved translated PPT");
            }
        } catch (Exception e) {
            log.error("Error during PPT translation:", e);
            throw e;
        }
    }
    
}
