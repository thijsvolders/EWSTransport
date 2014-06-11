package nl.yenlo.transport.msews.client;

import microsoft.exchange.webservices.data.BasePropertySet;
import microsoft.exchange.webservices.data.BodyType;
import microsoft.exchange.webservices.data.ConflictResolutionMode;
import microsoft.exchange.webservices.data.DeleteMode;
import microsoft.exchange.webservices.data.EmailMessage;
import microsoft.exchange.webservices.data.EmailMessageSchema;
import microsoft.exchange.webservices.data.ExchangeCredentials;
import microsoft.exchange.webservices.data.ExchangeService;
import microsoft.exchange.webservices.data.ExchangeVersion;
import microsoft.exchange.webservices.data.FindItemsResults;
import microsoft.exchange.webservices.data.FolderId;
import microsoft.exchange.webservices.data.ITraceListener;
import microsoft.exchange.webservices.data.Item;
import microsoft.exchange.webservices.data.ItemView;
import microsoft.exchange.webservices.data.PropertySet;
import microsoft.exchange.webservices.data.SearchFilter;
import microsoft.exchange.webservices.data.ServiceLocalException;
import microsoft.exchange.webservices.data.ServiceResponseException;
import microsoft.exchange.webservices.data.TraceFlags;
import microsoft.exchange.webservices.data.WebCredentials;
import microsoft.exchange.webservices.data.WellKnownFolderName;
import nl.yenlo.transport.msews.EWSPollTableEntry;
import nl.yenlo.transport.msews.client.exception.EwsMailClientCommunicationException;
import nl.yenlo.transport.msews.client.exception.EwsMailClientConfigException;
import nl.yenlo.transport.msews.log.EWSTraceListener;
import org.apache.commons.logging.Log;

import java.io.ByteArrayInputStream;
import java.io.InputStream;
import java.net.URI;
import java.net.URISyntaxException;
import java.util.EnumSet;
import java.util.Iterator;

/**
 * This is a class to connect with an Exchange WebServices enabled mailbox.
 * <p>
 * Please use as follows:
 * - Create an instance
 * - set login credentials (withLogin)
 * - set the serviceUrl (withServiceURL)
 * - set the folder to retrieve mails from
 * - (optional) set the batch size (is 10 per default)
 * - call getMailEntries to fetch the mails from the mailbox.
 * </p>
 *
 * @author tvolders.
 */
public class EwsMailClient {
    private Log log;
    private ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);

    private String username;
    private String password;
    private FolderId folder = null;
    private SearchFilter searchFilter = null;
    private PropertySet ewsProperties = new PropertySet(BasePropertySet.FirstClassProperties, EmailMessageSchema.Attachments);

    private int batchSize = 10;
    private FindItemsResults<Item> items = null;

    /**
     * Initialize the client.
     * <p>
     * Specify the log instance to log the clients actions to.
     * </p>
     *
     * @param log the logger to log the client's statements to
     */
    public EwsMailClient(Log log) {
        this.log = log;

        service.setTraceListener(new EWSTraceListener(log));

        if (log.isTraceEnabled()) {
            service.setTraceEnabled(true);
            service.setTraceFlags(EnumSet.allOf(TraceFlags.class));
        }

    }

    /**
     * Provide the login credentials for the mailbox to query
     *
     * @param emailAddress the mail address
     * @param password     the password associates with the account
     * @param domain       the domain (if required by the mailserver)
     * @return a client instance where the login credentials have been configured in
     */
    public EwsMailClient withLogin(String emailAddress, String password, String domain) {
        this.username = emailAddress;
        this.password = password;

        ExchangeCredentials credentials = new WebCredentials(emailAddress, password, domain);
        service.setCredentials(credentials);

        if (log.isDebugEnabled()) {
            log.debug("Using emailAddress " + emailAddress + " and a password");
        }

        return this;
    }

    /**
     * Supply the serviceURL of the EWS endpoint.
     * <p/>
     *
     * @param serviceURI the serviceURI of the EWS endpoint
     * @return the client instance with configured serviceURL
     * @see nl.yenlo.transport.msews.EWSPollTableEntry#getDomain()
     * </p>
     */
    public EwsMailClient withServiceURL(String serviceURI) {
        try {
            service.setUrl(new URI(serviceURI));
        } catch (URISyntaxException use) {
            throw new EwsMailClientConfigException("ServiceUrl appears to be in an incorrect format", use);
        }

        if (log.isDebugEnabled()) {
            log.debug("using MS Exchange Webservice url " + serviceURI);
        }

        return this;
    }

    /**
     * Perform autodiscovery of the configured mailbox.
     * <p>
     * When this fails its logged only. The exception is not rethrown.
     * </p>
     *
     * @return the client instance with autodiscovery performed (if possible)
     */
    public EwsMailClient withAutoDiscovery() {
        try {
            // Username IS the emailAddress
            if (log.isTraceEnabled()) {
                log.trace("Performing auto discovery for EWS-user: " + username);
            }
            service.autodiscoverUrl(username);
        } catch (Exception e) {
            // NOP....
            // Lets autodiscovery fail... Try whether the regular URL works...
            if (log.isInfoEnabled()) {
                log.info("Autodiscovery failed. Trying without autoDiscovery.");
            }

            if (log.isDebugEnabled()) {
                log.debug("Logging the AutoDiscovery error:", e);
            }
        }

        return this;
    }

    /**
     * OVerride the traceListener which is per default configured to log to the constructor-supplied log instance.
     *
     * @param traceListener the new tracelistener to log trace-events to
     */
    public void setTraceListener(ITraceListener traceListener) {
        service.setTraceListener(traceListener);
    }

    /**
     * The batchSize of the mail retrieval.
     *
     * @param batchSize the amount of mails to retrieve per polling-interval
     * @return the mailclient instance.
     */
    public EwsMailClient withBatchSize(int batchSize) {
        this.batchSize = batchSize;
        return this;
    }

    /**
     * Supply the folder to retrieve mails from.
     *
     * @param folder a WellKnownFolderName
     * @return the mailClient instance.
     */
    public EwsMailClient forFolder(FolderId folder) {
        this.folder = folder;
        return this;
    }

    public EwsMailClient withSearchFilter(SearchFilter searchFilter) {
        this.searchFilter = searchFilter;
        return this;
    }

    /**
     * Get the mailEntries
     * <p>
     * This method will invoke the EWS services to retrieve the mails per the provided configuration.
     * </p>
     * <p>
     * The configuration should have been provided through the various fluent-api methods.
     * </p>
     */
    public void getMailEntries() {
        try {
            ItemView iv = new ItemView(batchSize);

            if (log.isTraceEnabled()) {
                log.trace("Finding items in the mail-folder " + folder.getUniqueId());
            }

            items = service.findItems(folder, searchFilter, iv);
            if (log.isDebugEnabled()) {
                log.debug("Retrieved " + items.getTotalCount() + " messages from mailbox");
            }
            if (log.isTraceEnabled()) {
                log.trace("Loading item properties");
            }

            if (items != null && items.getTotalCount() > 0) { // Check whether there are items before loading the properties
                service.loadPropertiesForItems(items, ewsProperties);
            }

        } catch (Exception e) {
            throw new EwsMailClientCommunicationException("A communication exception occurred while retrieving mail items", e);
        }
    }


    /**
     * Return an iterator for the mailEntries.
     * <p>
     * When the mailEntries have not been fetched yet they will be retrieved automatically.
     * </p>
     *
     * @return an iterator over the retrieved mailEntries.
     */
    public Iterator<EmailMessage> getMailEntryIterator() {
        Iterator<EmailMessage> result = null;
        if (this.items == null) {
            // Get the mailItems first..
            getMailEntries();
        }

        if (this.items != null) {
            // Then create an iterator and return a self-progressing iterator... We hide the FindItemResult...
            result = new Iterator<EmailMessage>() {
                private Iterator<Item> internal = items.iterator();

                @Override
                public boolean hasNext() {
                    return internal.hasNext();
                }

                /**
                 * Get the emailmessage from the items collection, load its properties, create the mailmessage instance and return it.
                 * @return the prepared emailmessage
                 */
                @Override
                public EmailMessage next() {
                    EmailMessage message = null;

                    try {
                        Item next = internal.next();
                        if (log.isTraceEnabled()) {
                            String itemId = "unknown";
                            if (next.getId() != null) {
                                itemId = next.getId().getUniqueId();
                            }
                            log.trace("Loading item '" + itemId + "' itself");
                        }

                        // Source: http://stackoverflow.com/a/21772997
                        next.load(ewsProperties);

                        // Bind to an existing message using its unique identifier.
                        message = EmailMessage.bind(service, next.getId());

                        if (log.isInfoEnabled()) {
                            log.info("Loaded email '" + message.getSubject() + "' sent from '" + message.getSender().toString() + "'");
                        }

                    } catch (Exception e) {
                        throw new EwsMailClientCommunicationException("Loading item has failed", e);
                    }

                    return message;
                }

                @Override
                public void remove() {
                    internal.remove();
                }
            };
        }

        return result;
    }

    /**
     * Get the mail's body as InputStream.
     *
     * @param message the mailmessage to get the body for.
     * @return an inputstream with the content of the mail body.
     */
    public InputStream getBodyAsInputStream(final EmailMessage message) {
        if (log.isTraceEnabled()) {
            log.trace("The mail has NO attachments. Using the body as message.");
        }
        try {
            if (message.getBody().getBodyType() == BodyType.HTML) {
                throw new RuntimeException("HTML bodytypes are not supported!!");
            }

            // Get the body text and put that into the InputStream
            return new ByteArrayInputStream(message.getBody().toString().getBytes());
        } catch (ServiceLocalException sle) {
            throw new EwsMailClientCommunicationException("Could not extract body from the email.", sle);
        }
    }

    /**
     * Retrieve the contentType of the mailBody from the supplied message
     *
     * @param message the message to get the contentType from..
     * @return the contentType
     */
    public String getBodyContentType(final EmailMessage message) {
        return "application/xml"; // default is XML...
    }

    /**
     * Delete a message from the mailbox
     *
     * @param message the message to delete
     */
    public void deleteMessage(EmailMessage message, EWSPollTableEntry.DeleteActionType deleteActionType) {
        // Be carefull here!!! You will be deleting mails from the mailbox!!
        try {
            log.info("Message to indicate that a mail would be deleted with subject : " + message.getSubject());
            //message.delete(deleteActionType == EWSPollTableEntry.DeleteActionType.TRASH ? DeleteMode.SoftDelete : DeleteMode.HardDelete);
        } catch (Exception e) {
            log.error("Could not successfully delete the message. ", e);
            throw new EwsMailClientCommunicationException("Could not successfully delete the message. ", e);
        }
    }

    /**
     * Move the message to another folder
     *
     * @param message  the message to move
     * @param toFolder the folder to move it to (can be a WellKnownFolderName)
     */
    public void moveMessage(EmailMessage message, String toFolder) {
        try {
            // We will move the message.
            message.move(new FolderId(toFolder));
        } catch (Exception e) {
            throw new EwsMailClientCommunicationException("A communication exception occurred while moving the email message", e);
        }
    }

    public void markAsRead(EmailMessage message) {
        try {
            message.setIsRead(true);
            message.update(ConflictResolutionMode.AutoResolve);
        } catch (ServiceResponseException sre) {
            throw new EwsMailClientCommunicationException("Could not mark message as 'read'. ", sre);
        } catch (Exception e) {
            throw new EwsMailClientCommunicationException("Could not mark message as 'read'. ", e);
        }

    }

}
