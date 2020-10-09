package ru.ik87;

import java.io.FileOutputStream;
import java.io.IOException;

public class SaveLocal {
    public void save(byte[] is, String emailTo) throws IOException {
        try (FileOutputStream fos = new FileOutputStream("out/"+emailTo+".docx")) {
            fos.write(is);
        }
    }
}
