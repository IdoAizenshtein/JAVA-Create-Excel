package bizibox.exportsExcelMail;

import javax.mail.*;
import javax.mail.internet.*;
import java.io.File;
import java.util.Date;
import java.util.Properties;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeMessage;

public class SendFileEmail {
    final String username = "alerts@bizibox.biz";
    final String password = "bizi2012";
    Session session;

    SendFileEmail() {
        Properties props = new Properties();
        props.put("mail.smtp.host", "smtp.gmail.com");
        props.put("mail.smtp.socketFactory.port", "465");
        props.put("mail.smtp.socketFactory.class", "javax.net.ssl.SSLSocketFactory");
        props.put("mail.smtp.auth", "true");
        props.put("mail.smtp.port", "465");
        props.put("mail.smtp.ssl.enable", "true");

        this.session = Session.getDefaultInstance(props,
                new Authenticator() {
                    protected PasswordAuthentication getPasswordAuthentication() {
                        return new PasswordAuthentication(username, password);
                    }
                });
    }

    public String sender(File fileXlsx, String title, String name_company, String name_roh, String name_doch, String toAddressMail) throws MessagingException, Exception {
        try {
            // creates a new e-mail message
            MimeMessage message = new MimeMessage(this.session);

            message.setFrom(new InternetAddress("bizibox <info@bizibox.biz>"));
            InternetAddress[] toAddresses = {new InternetAddress(toAddressMail)};
            message.setRecipients(Message.RecipientType.TO, toAddresses);
            message.setSubject(title, "UTF-8");
            message.setSentDate(new Date());

            //creates message part
            MimeBodyPart messageBodyPart = new MimeBodyPart();
            messageBodyPart.setContent(
                    "<html><body style=\"margin:0;padding: 34px 0px 0px 0px;font-size: 16px;background-color: #e8e8e8;color: #222222;direction: rtl;\"><table cellpadding=\"0\" cellspacing=\"0\" border=\"0\" width=\"100%\" align=\"center\" class=\"bgBody\"><tbody><tr><td align=\"center\"><table cellpadding=\"0\" cellspacing=\"0\" width=\"701\" border=\"0\" style=\"background: #fff;\"><tbody><tr></tr><tr><td align=\"center\"><img src=\"https://secure.bizibox.biz/newslleter/header.png\" data-default=\"placeholder\" width=\"701\" height=\"343\" data-max-width=\"701\"></td></tr><tr><td class=\"bgItem\" align=\"center\" width=\"600\" bgcolor=\"#fff\"><table cellpadding=\"0\" cellspacing=\"0\" width=\"540\" border=\"0\"><tbody><tr><td class=\"movableContentContainer\"><table cellpadding=\"0\" cellspacing=\"0\" border=\"0\" width=\"570\"><tbody><tr><td height=\"44\"></td></tr><tr><td><h2 style=\"color: #333133 !important;font: bold 26px/35px arial;padding: 0;margin: 0;direction: rtl;text-align: right;display: block;\"> היי " +
                            name_company +
                            ",</h2><p style=\"text-align: right;direction:rtl;font: 16px/20px arial;color: #4e4c4e;padding: 0;margin: 0;\">" +
                            name_roh +
                            " שלח לך דו\"ח" +
                            name_doch +
                            ".</p><p style=\"text-align:right;direction:rtl;font:16px/20px arial;color:#4e4c4e;padding:0;margin:0\">שנערך באמצעות תוכנת <a style=\"color:#1387a9\" href=\"http://bizibox.biz/\" target=\"_blank\">bizibox</a>.</p></td></tr></tbody></table></td></tr>" +
                            "<td align=\"center\" style=\" height: 170px;\"><img src=\"https://secure.bizibox.biz/newslleter/logo.png\" data-default=\"placeholder\" width=\"223\" height=\"56\" data-max-width=\"223\"></td></tbody></table></td></tr><tr><td><table cellpadding=\"0\" cellspacing=\"0\" border=\"0\" style=\"background:#e8e8e8;text-align: center;vertical-align: middle;\" width=\"701\"><tbody><tr><td colspan=\"4\" height=\"15\"></td></tr><tr><td valign=\"middle\" style=\"text-align:center\"><td valign=\"middle\" style=\"text-align:center\"><p style=\"direction:rtl;font:12px arial;text-align:center;vertical-align:middle;display:inline-block;color:#4e4c4e;width:570px\"> במידה ואין ברצונך לקבל אימיילים נוספים מביזיבוקס בעתיד, ניתן <a style=\"color:#2284a1;text-decoration:underline\">לבטל את המנוי</a> למידע נוסף לגבי ביזיבוקס ניתן להיכנס לאתר החברה בכתובת <a href=\"http://bizibox.biz/\" style=\"color:#2284a1;text-decoration:underline\" target=\"_blank\">www.bizibox.biz</a></p></td></td></tr><tr><td colspan=\"4\" height=\"15\"></td></tr></tbody></table></td></tr></tbody></table></td></tr></tbody></table></body></html>",
                    "text/html; charset=UTF-8");

            // creates multi-part
            Multipart multipart = new MimeMultipart();
            multipart.addBodyPart(messageBodyPart);
            MimeBodyPart attachPart = new MimeBodyPart();

            attachPart.attachFile(fileXlsx);
            attachPart.setHeader("Content-Type", "application/vnd.ms-excel");

            multipart.addBodyPart(attachPart);
            message.setContent(multipart);

            // sends the e-mail
            Transport.send(message);
            return "true";
        } catch (MessagingException e) {
            String err = "MessagingException thrown  :" + e;
            return err;
        } catch (Exception e) {
            String err = "Exception thrown  :" + e;
            return err;
        }
    }
}


