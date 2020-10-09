package ru.ik87;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import ru.ik87.smtp.Email;
import ru.ik87.xwpf.TableXlsx;
import ru.ik87.xwpf.ReplaceText;

import java.io.IOException;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.ThreadPoolExecutor;

public class Program {
    private static final Logger LOGGER = LogManager.getLogger(Program.class.getName());

    public static void main(String[] args) {
        Config mainConfig = new Config("main.properties");
        Config emailConfig = new Config("email.properties");
        String tableXlsxFIle = mainConfig.getProperties().getProperty("tableXlsxFIle");
        String templateDocxFIle = mainConfig.getProperties().getProperty("templateDocxFIle");
        ExecutorService pool = Executors.newFixedThreadPool(Runtime.getRuntime().availableProcessors());
        SaveLocal saveLocal = new SaveLocal();
        String state = mainConfig.getProperties().getProperty("state");
        Email email = new Email(emailConfig.getProperties());
        String email_col = emailConfig.getProperties().getProperty("email");
        TableXlsx tableXlsx = new TableXlsx(tableXlsxFIle);
        ReplaceText replaceText = new ReplaceText(templateDocxFIle);
        System.out.println("Начало отправки");
        try {
            List<Map<String, String>> tables = tableXlsx.readXlsx();
            int i = 1;
            for (Map<String, String> row : tables) {
                if (!row.isEmpty()) {
                    byte[] doc = replaceText.composeDocx(row);
                    if (row.get(state) == null) {
                       // email.send(doc, row.get(email_col));
                        saveLocal.save(doc, row.get(email_col) );
                        System.out.println(row.get(email_col) + " отправленно");
                        tableXlsx.writeXlsx(i, state, "отправленно");
                    }
                }
                i++;
            }
        } catch (IOException | InvalidFormatException /*| EmailException */e) {
            LOGGER.error(e.getMessage());
        }
        System.out.println("Отправка завершена");
    }
}
