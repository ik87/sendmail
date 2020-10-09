package ru.ik87.xwpf;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.*;

import java.io.*;
import java.util.List;
import java.util.Map;

public class ReplaceText {
    private final String template;

    public ReplaceText(String templateDocxFile) {
        this.template = templateDocxFile;
    }

    public byte[] composeDocx(Map<String, String> variables) throws IOException {
        FileInputStream fis = new FileInputStream(new File(template));
        XWPFDocument doc = new XWPFDocument(fis);
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

        fis.close();
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
