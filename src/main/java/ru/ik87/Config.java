package ru.ik87;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.nio.charset.Charset;
import java.util.Properties;

/**
 * @author Kosolapov Ilya (d_dexter@mail.ru)
 * @version $Id$
 * @since 0.1
 */
public class Config {
    private final Properties values = new Properties();

    public  Config(String fileConfig) {
        try (InputStream in  = new FileInputStream(new File(fileConfig))) {
            values.load(new InputStreamReader(in, Charset.forName("UTF-8")));
        } catch (Exception e) {
            throw new IllegalStateException(e);
        }
    }

    public Properties getProperties() {
        return values;
    }
}
