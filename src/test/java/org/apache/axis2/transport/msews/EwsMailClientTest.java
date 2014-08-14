package org.apache.axis2.transport.msews;

import microsoft.exchange.webservices.data.EmailMessage;
import microsoft.exchange.webservices.data.EmailMessageSchema;
import microsoft.exchange.webservices.data.FolderId;
import microsoft.exchange.webservices.data.LogicalOperator;
import microsoft.exchange.webservices.data.SearchFilter;
import microsoft.exchange.webservices.data.WellKnownFolderName;
import nl.yenlo.transport.msews.client.EwsMailClient;
import org.apache.axis2.util.IOUtils;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.junit.Before;
import org.junit.Ignore;
import org.junit.Test;

import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;

/**
 * A simple yet helpful testclass.
 * This class cannot be used automatically but should be executed manually. Therefor it is ignored through annotations.
 *
 * @author tvolders.
 */
@Ignore
public class EwsMailClientTest {
    private final static Log logger = LogFactory.getLog(EwsMailClientTest.class);
    public static final String DOMAIN = "nl";
    // System properties
    // password
    public static final String PW_SYSTEM_PROP = "pw";
    // username
    public static final String UN_SYSTEM_PROP = "un";
    // domain
    public static final String DOMAIN_SYSTEM_PROP = "domain";
    // service url, something like "https://xx.xx.xx.xx/EWS/Exchange.asmx";
    public static final String SU_SYSTEM_PROP = "su"; //serviceURL

    private EwsMailClient client;

    @Before
    public void setup() throws IOException {
        client = new EwsMailClient(logger);

        String pw = System.getProperty(PW_SYSTEM_PROP);
        String un = System.getProperty(UN_SYSTEM_PROP);
        String domain = System.getProperty(DOMAIN_SYSTEM_PROP);
        String serviceUrl = System.getProperty(SU_SYSTEM_PROP);
        client.withLogin(un, pw, DOMAIN).withServiceURL(serviceUrl);

        client.forFolder(new FolderId(WellKnownFolderName.Inbox).toString());
        SearchFilter.SearchFilterCollection sf = new SearchFilter.SearchFilterCollection(LogicalOperator.And,
                new SearchFilter.ContainsSubstring(EmailMessageSchema.Sender, "@" + domain), new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false) /*,
                new SearchFilter.ContainsSubstring(EmailMessageSchema.Subject, "Test")*/);

        client.withSearchFilter(sf);
        client.withBatchSize(3).getMailEntries();

    }

    @Test
    public void test() throws Exception {
        Iterator<EmailMessage> mailEntryIterator = client.getMailEntryIterator();

        while (mailEntryIterator.hasNext()) {
            EmailMessage item = mailEntryIterator.next();

            item.getId();

            InputStream bodyAsInputStream = client.getBodyAsInputStream(item);

            IOUtils.copy(bodyAsInputStream, System.out, false);

           // client.deleteMessage(item, PollTableEntry.DeleteActionType.TRASH);
            client.markAsRead(item);

        }
    }
}
