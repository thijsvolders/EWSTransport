package org.apache.axis2.transport.msews;

import microsoft.exchange.webservices.data.EmailMessage;
import microsoft.exchange.webservices.data.EmailMessageSchema;
import microsoft.exchange.webservices.data.LogicalOperator;
import microsoft.exchange.webservices.data.SearchFilter;
import microsoft.exchange.webservices.data.WellKnownFolderName;
import nl.yenlo.transport.msews.PollTableEntry;
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
    public static final String SERVICE_URL = "YourServerHere";
    public static final String USERACCOUNT = "YourAccountHere";
    public static final String DOMAIN = "YouDomainHere";
    public static final String PW_SYSTEM_PROP = "pw";
    private EwsMailClient client;

    @Before
    public void setup() throws IOException {
        client = new EwsMailClient(logger);

        String pw = System.getProperty(PW_SYSTEM_PROP);
        client.withLogin(USERACCOUNT, pw, DOMAIN).withServiceURL(SERVICE_URL);

        client.forFolder(WellKnownFolderName.Inbox);
        SearchFilter.SearchFilterCollection sf = new SearchFilter.SearchFilterCollection(LogicalOperator.And,
                new SearchFilter.ContainsSubstring(EmailMessageSchema.Sender, "@example.com"),
                new SearchFilter.ContainsSubstring(EmailMessageSchema.Subject, "Test"));

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

        }
    }
}
