package ru.ik87.xwpf;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.*;

import java.io.*;
import java.nio.file.Files;

import java.util.List;
import java.util.Map;

public class ReplaceText {
    private final byte[] template;

    public ReplaceText(String templateDocxFile) throws IOException {
        File file = new File(templateDocxFile);
        this.template =  Files.readAllBytes(file.toPath());
    }

    public byte[] generateFromTemplate(Map<String, String> variables) throws IOException {
        InputStream bis = new ByteArrayInputStream(template);
        XWPFDocument doc = new XWPFDocument(bis);
        for (XWPFParagraph p : doc.getParagraphs()) {
            List<XWPFRun> runs = p.getRuns();
            if (runs != null) {
                for (XWPFRun r : runs) {
                    replace(r, variables);
                }
            }
        }
        for (XWPFTable tbl : doc.getTables()) {
            for (XWPFTableRow row : tbl.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (XWPFParagraph p : cell.getParagraphs()) {
                        for (XWPFRun r : p.getRuns()) {
                            replace(r, variables);
                        }
                    }
                }
            }
        }

        bis.close();
        ByteArrayOutputStream buffer = new ByteArrayOutputStream();
        doc.write(buffer);
        byte[] bytes = buffer.toByteArray();
        buffer.close();
        doc.close();

        return bytes;
    }


    private void replace(XWPFRun r, Map<String, String> variables) {
        String text = r.getText(0);
        if (text != null) {
            for (var variable : variables.entrySet()) {
                if (text.contains(variable.getKey())) {
                    text = text.replace(variable.getKey(), variable.getValue());
                    r.setText(text, 0);
                }
            }
        }
    }


}
