package ru.ik87.smtp;

import org.apache.commons.mail.*;

import javax.mail.util.ByteArrayDataSource;
import java.util.Properties;

public class Email {
    private HtmlEmail email;
    private final String emailAttachFile;
    private final String emailAttachDecrFile;
    private final Integer smtp;
    private final String emailFrom;
    private final String hostName;
    private final String emailLogin;
    private final String emailPass;
    private final String from_name;
    private final String subject;
    private final String html_msg;
    private final String text_msg;

    public Email(Properties config) {
        this.emailAttachFile = config.getProperty("email_attach_file");
        this.emailAttachDecrFile = config.getProperty("email_attach_descr_file");
        this.smtp = Integer.valueOf(config.getProperty("email_smtp_port", "2525"));
        this.emailFrom = config.getProperty("email_from");
        this.hostName = config.getProperty("email_host_name");
        this.emailLogin = config.getProperty("email_auth_login");
        this.emailPass = config.getProperty("email_auth_pass");
        this.from_name = config.getProperty("email_from_name");
        this.subject = config.getProperty("email_subject");
        this.html_msg = config.getProperty("email_html_msg");
        this.text_msg = config.getProperty("email_text_msg");
    }

    public void init() throws EmailException {
        this.email = new HtmlEmail();
        this.email.setSmtpPort(smtp);
        this.email.setFrom(emailFrom, from_name);
        this.email.setHostName(hostName);
        this.email.setAuthentication(emailLogin, emailPass);
        this.email.setSubject(subject);
        this.email.setHtmlMsg(html_msg);
        this.email.setTextMsg(text_msg);
        this.email.setCharset("utf-8");
        this.email.setSSLOnConnect(false);

    }

    public void send(byte[] is, String emailTo) throws EmailException {
        init();
        this.email.addTo(emailTo);
        this.email.addReplyTo(emailFrom, "some@mail.ru");
        var source = new ByteArrayDataSource(is, "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
        //  add the attachment
        this.email.attach(source, emailAttachFile, emailAttachDecrFile);
        // send the email
        email.send();
    }
}
