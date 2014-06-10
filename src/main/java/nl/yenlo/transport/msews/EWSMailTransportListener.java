/*
 *  Licensed to the Apache Software Foundation (ASF) under one
 *  or more contributor license agreements.  See the NOTICE file
 *  distributed with this work for additional information
 *  regarding copyright ownership.  The ASF licenses this file
 *  to you under the Apache License, Version 2.0 (the
 *  "License"); you may not use this file except in compliance
 *  with the License.  You may obtain a copy of the License at
 *
 *   http://www.apache.org/licenses/LICENSE-2.0
 *
 *  Unless required by applicable law or agreed to in writing,
 *  software distributed under the License is distributed on an
 *   * "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
 *  KIND, either express or implied.  See the License for the
 *  specific language governing permissions and limitations
 *  under the License.
 */

package nl.yenlo.transport.msews;

import microsoft.exchange.webservices.data.AffectedTaskOccurrence;
import microsoft.exchange.webservices.data.Attachment;
import microsoft.exchange.webservices.data.AttachmentCollection;
import microsoft.exchange.webservices.data.BodyType;
import microsoft.exchange.webservices.data.DeleteMode;
import microsoft.exchange.webservices.data.EmailAddress;
import microsoft.exchange.webservices.data.EmailAddressCollection;
import microsoft.exchange.webservices.data.EmailMessage;
import microsoft.exchange.webservices.data.ExchangeService;
import microsoft.exchange.webservices.data.FileAttachment;
import microsoft.exchange.webservices.data.InternetMessageHeader;
import microsoft.exchange.webservices.data.InternetMessageHeaderCollection;
import microsoft.exchange.webservices.data.ItemId;
import microsoft.exchange.webservices.data.SendCancellationsMode;
import microsoft.exchange.webservices.data.ServiceLocalException;
import org.apache.axis2.AxisFault;
import org.apache.axis2.Constants;
import org.apache.axis2.addressing.EndpointReference;
import org.apache.axis2.context.MessageContext;
import org.apache.axis2.description.TransportInDescription;
import org.apache.axis2.transport.TransportUtils;
import org.apache.axis2.transport.base.AbstractPollingTransportListener;
import org.apache.axis2.transport.base.BaseConstants;
import org.apache.axis2.transport.base.ManagementSupport;
import org.apache.axis2.transport.base.event.TransportErrorListener;
import org.apache.axis2.transport.base.event.TransportErrorSource;
import org.apache.axis2.transport.base.event.TransportErrorSourceSupport;
import nl.yenlo.transport.msews.client.EwsMailClient;

import javax.mail.MessagingException;
import javax.mail.internet.AddressException;
import javax.mail.internet.InternetAddress;
import javax.xml.stream.XMLStreamException;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.PipedInputStream;
import java.io.PipedOutputStream;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;
import java.util.concurrent.ConcurrentHashMap;

/**
 * This mail transport lister implementation uses the base transport framework and is a polling
 * transport. i.e. a service can register itself with custom a custom mail configuration (i.e.
 * pop3 or imap) and specify its polling duration, and what action to be taken after processing
 * messages. The transport always deletes processed mails from the folder they were fetched from
 * and can be configured to be optionally moved to a different folder, if the server supports it
 * (e.g. with imap). When checking for new mail, the transport ignores messages already flaged as
 * SEEN and DELETED
 */

public class EWSMailTransportListener extends AbstractPollingTransportListener<PollTableEntry>
        implements ManagementSupport, TransportErrorSource {

    private final TransportErrorSourceSupport tess = new TransportErrorSourceSupport(this);
    private static String transportName = null;

    @Override
    protected void doInit() throws AxisFault {
        super.doInit();

        // Lets find ourselves and get the name of the transport. Its used in the endpoint prefix...
        HashMap<String, TransportInDescription> transportsIn = cfgCtx.getAxisConfiguration().getTransportsIn();
        for (Map.Entry<String, TransportInDescription> tid : transportsIn.entrySet()) {
            // Find ourselves... (the class)
            if (tid.getValue().getReceiver() instanceof EWSMailTransportListener) {
                // Found it...
                transportName = tid.getKey();
            }
        }

        // set the synchronise callback table
        if (cfgCtx.getProperty(BaseConstants.CALLBACK_TABLE) == null) {
            cfgCtx.setProperty(BaseConstants.CALLBACK_TABLE, new ConcurrentHashMap());
        }

        log.info("Initializing Exchange WS 2013 Listener (" + transportName + ")...");
    }

    @Override
    protected void poll(PollTableEntry entry) {
        try {
            checkMail(entry, entry.getEmailAddress());
        } catch (Exception e) {
            // A catch all construction where we can log any exception which was uncaughtin the checkMail method
            processFailure("An unexpected error occurred while polling the EWS-mail server", e, entry);
        }
    }

    /**
     * Check mail for a particular service that has registered with the mail transport
     *
     * @param entry        the poll table entry that stores service specific informaiton
     * @param emailAddress the email address checked
     */
    private void checkMail(final PollTableEntry entry, InternetAddress emailAddress) {

        if (log.isDebugEnabled()) {
            log.debug("Checking mail for account : " + emailAddress);
        }

        try {
            if (log.isDebugEnabled()) {
                log.debug("Attempting to connect to EWS server (" + entry.getServiceUrl() + ") for : " + entry.getEmailAddress());
            }

            EwsMailClient client = new EwsMailClient(log);
            if (entry.getEmailAddress() != null && entry.getPassword() != null) {
                client.withLogin(entry.getEmailAddress().getAddress(), entry.getPassword(), entry.getDomain()).withServiceURL(entry.getServiceUrl()).withAutoDiscovery();
            } else {
                throw new RuntimeException("Unable to locate username and/or password for mail login");
            }

            client.withBatchSize(entry.getMessageCount()).getMailEntries();

            client.forFolder(entry.getFolder());

            Iterator<EmailMessage> mailEntryIterator = client.getMailEntryIterator();

            outer:
            while (mailEntryIterator.hasNext()) {
                final EmailMessage item = mailEntryIterator.next();

                entry.processingUID(item.getId().getUniqueId());

                Runnable onCompletion = new MailCheckCompletionTask(emailAddress, entry);

                if (log.isTraceEnabled()) {
                    log.trace("Binding item " + item.getId().getUniqueId() + " to an emailMessage instance");
                }

                if (item != null) {   // Not sure whether message CAN be null
                    if (log.isTraceEnabled()) {
                        log.trace("processing the message");
                    }
                    processMail(entry, item, onCompletion, client);
                } else {
                    if (log.isTraceEnabled()) {
                        log.trace("mesage is null, running onCompletion");
                    }
                    onCompletion.run();
                }
            }
        } catch (Exception sle) {
            throw new RuntimeException("An error occurred while communicating with the Exchange Webservices", sle);

        }
    }

    /**
     * Invoke the actual message processor in the current thread or another worker thread
     *
     * @param entry        PolltableEntry
     * @param message      message to process
     * @param onCompletion the tasks to run on the completion of mail processing
     */
    private void processMail(PollTableEntry entry, EmailMessage message, Runnable onCompletion, EwsMailClient client) throws ServiceLocalException {

        MailProcessor mp = new MailProcessor(entry, message, client, onCompletion);
        String msgId = message.getId().getUniqueId();

        // should messages be processed in parallel?
        if (entry.isConcurrentPollingAllowed()) {

            // try to locate the UID of the message
            String uid = getMessageUID(message);

            if (uid != null) {
                if (entry.isProcessingUID(uid)) {
                    if (log.isDebugEnabled()) {
                        log.debug("Skipping message # : " + msgId + " : UIDL " + uid + " - already being processed by another thread");
                    }
                } else {
                    mp.setUID(uid);

                    if (entry.isProcessingMailInParallel()) {
                        if (log.isDebugEnabled()) {
                            log.debug("Processing message # : " + msgId + " with UID : " + uid + " with a worker thread");
                        }
                        workerPool.execute(mp);
                    } else {
                        if (log.isDebugEnabled()) {
                            log.debug("Processing message # : " + msgId + " with UID : " + uid + " in same thread");
                        }
                        mp.run();
                    }
                }
            } else {
                log.warn("Cannot process mail in parallel as the " + "folder does not support UIDs. Processing message # : " + msgId + " in the same thread");
                entry.setConcurrentPollingAllowed(false);
                mp.run();
            }

        } else {
            if (entry.isProcessingMailInParallel()) {
                if (log.isDebugEnabled()) {
                    log.debug("Processing message # : " + msgId +
                            " with a worker thread");
                }
                workerPool.execute(mp);
            } else {
                if (log.isDebugEnabled()) {
                    log.debug("Processing message # : " + msgId + " in same thread");
                }
                mp.run();
            }
        }
    }

    /**
     * Handle processing of a message, possibly in a new thread
     */
    private class MailProcessor implements Runnable {

        private PollTableEntry entry = null;
        private EmailMessage message = null;
        private String uid = null;
        private Runnable onCompletion = null;
        private EwsMailClient client = null;

        MailProcessor(PollTableEntry entry, EmailMessage message, final EwsMailClient client, Runnable onCompletion) {
            this.entry = entry;
            this.message = message;
            this.onCompletion = onCompletion;

            this.client = client;
        }

        public void setUID(String uid) {
            this.uid = uid;
        }

        public void run() {

            entry.setLastPollState(PollTableEntry.NONE);
            try {
                processMail(message, entry, client);
                entry.setLastPollState(PollTableEntry.SUCCSESSFUL);
                metrics.incrementMessagesReceived();

            } catch (Exception e) {
                log.error("Failed to process message", e);
                entry.setLastPollState(PollTableEntry.FAILED);
                metrics.incrementFaultsReceiving();
                tess.error(entry.getService(), e);

            } finally {
                if (uid != null) {
                    entry.removeUID(uid);
                }
            }
            try {
                moveOrDeleteAfterProcessing(entry, client, message);
            } catch (Exception e) {
                log.error("Failed to move or delete email message", e);
                tess.error(entry.getService(), e);
            }

            // Old code counted towards 0 then at the end of the proces ran this oncompletion.run method
            onCompletion.run();
        }
    }

    /**
     * Handle optional logic of the mail transport, that needs to happen once all messages in
     * a check mail cycle has ended.
     */
    private class MailCheckCompletionTask implements Runnable {
        private final InternetAddress emailAddress;
        private final PollTableEntry entry;
        private boolean taskStarted = false;

        public MailCheckCompletionTask(InternetAddress emailAddress, PollTableEntry entry) {
            this.emailAddress = emailAddress;
            this.entry = entry;
        }

        public void run() {
            synchronized (this) {
                if (taskStarted) {
                    return;
                } else {
                    taskStarted = true;
                }
            }

            if (log.isDebugEnabled()) {
                log.debug("Executing onCompletion task for the mail download of : " + emailAddress);
            }

            if (log.isDebugEnabled()) {
                log.debug("Scheduling next poll for : " + emailAddress);
            }
            onPollCompletion(entry);
        }
    }

    /**
     * Process a mail message through Axis2
     *
     * @param message the email message
     * @param entry   the poll table entry
     * @throws MessagingException on error
     * @throws IOException        on error
     */
    private void processMail(EmailMessage message, PollTableEntry entry, EwsMailClient client)
            throws Exception {

        if (log.isDebugEnabled()) {
            log.debug("Processing message with subject: '" + message.getSubject() + "' from '" + message.getFrom().getAddress() + "'.");
        }

        updateMetrics(message);

        // populate transport headers using the mail headers
        Map trpHeaders = getTransportHeaders(message);

        // set the message payload to the message context
        InputStream inputStream = null;

        MessageContext msgContext = entry.createMessageContext();

        MailOutTransportInfo outInfo = buildOutTransportInfo(message, entry);

        // save out transport information
        msgContext.setProperty(Constants.OUT_TRANSPORT_INFO, outInfo);

        // set message context From
        if (outInfo.getFromAddress() != null) {
            msgContext.setFrom(new EndpointReference(transportName + outInfo.getFromAddress().getAddress()));
        }

        // save original mail message id message context MessageID
        msgContext.setMessageID(outInfo.getRequestMessageID());

        String contentType = null;

        // If set to process the attachments, go
        // else, use the message body as MEssagecontext SoapEnvelope..
        if (entry.getExtractType() == PollTableEntry.ExtractType.BODY) {
            inputStream = client.getBodyAsInputStream(message);
        } else {
            // Untested code!!!

            FileAttachment fa = null;
            if (message.getHasAttachments()) {
                if (log.isTraceEnabled()) {
                    log.trace("The mail has attachments");
                }

                // We must have an attachment
                // FIXME: Check against regex whether this is interesting
                AttachmentCollection attachments = message.getAttachments();

                inputStream = new PipedInputStream();
                PipedOutputStream pop = new PipedOutputStream((PipedInputStream) inputStream);

                for (Attachment attachment : attachments) {
                    // LOG the attachment info

                    if (attachment instanceof FileAttachment) {
                        fa = (FileAttachment) attachment;

                        try {
                            // Only load the FileAttachment when its present
                            if (fa != null) {
                                fa.load(pop);
                            }
                        } catch (IOException ioe) {
                            throw new RuntimeException("An error occurred loading the file attachment for message : " + message.getId());
                        }

                        contentType = fa.getContentType();

                        if (log.isTraceEnabled()) {
                            log.trace("Going to create a SOAP Envelope...");
                        }

                        //Step out of this for loop
                        break;
                    } else {
                        if (log.isInfoEnabled()) {
                            log.info("An attachment of an unknown type has been discovered (type found is : " + attachment.getClass().getName() + ")");
                        }
                        // LOG strange attachment and throw away of move....
                    }
                }
            }
        }

        if (log.isTraceEnabled()) {
            log.trace("Constructing stream for attachment handling...");
        }

        // When there are no attachments then set the message body as XML content.
        try {
            try {
                msgContext.setEnvelope(TransportUtils.createSOAPMessage(msgContext, inputStream, contentType));
            } catch (XMLStreamException ex) {
                handleException("Error parsing message", ex);
            }

            String soapAction = (String) trpHeaders.get(BaseConstants.SOAPACTION);

            // Allow the subject to define the required soapAction on the destination message.
            if (soapAction == null && message.getSubject() != null &&
                    message.getSubject().startsWith(BaseConstants.SOAPACTION)) {
                soapAction = message.getSubject().substring(BaseConstants.SOAPACTION.length());
                if (soapAction.startsWith(":")) {
                    soapAction = soapAction.substring(1).trim();
                }
            }

            handleIncomingMessage(msgContext, trpHeaders, soapAction, contentType);
        } finally {
            try {
                inputStream.close();
            } catch (Exception e) {
                // LOG this!! , Do not break the execution
                log.error("An exception occurred while closing the inputstream", e);
            }
        }

        if (log.isDebugEnabled()) {
            log.debug("Processed message : " + message.getInternetMessageId() + " :: " + message.getSubject());
        }
    }

    private void updateMetrics(EmailMessage message) throws ServiceLocalException {
        int size = message.getSize();
        if (size != -1) {
            metrics.incrementBytesReceived(size);
        }
    }

    private Map getTransportHeaders(EmailMessage message) throws ServiceLocalException {

        //use a comaprator to ignore the case for headers.
        Comparator comparator = new Comparator() {
            public int compare(Object o1, Object o2) {
                String string1 = (String) o1;
                String string2 = (String) o2;
                return string1.compareToIgnoreCase(string2);
            }
        };

        final Map trpHeaders = new TreeMap(comparator);

        InternetMessageHeaderCollection internetMessageHeaders = message.getInternetMessageHeaders();
        for (InternetMessageHeader internetMessageHeader : internetMessageHeaders) {
            trpHeaders.put(internetMessageHeader.getName(), internetMessageHeader.getValue());
        }

        return trpHeaders;
    }

    private MailOutTransportInfo buildOutTransportInfo(EmailMessage message,
                                                       PollTableEntry entry) throws ServiceLocalException, AddressException {
        MailOutTransportInfo outInfo = new MailOutTransportInfo(entry.getEmailAddress());

        // determine reply address
        EmailAddressCollection replyTo = message.getReplyTo();
        if (replyTo != null) {
            final List<InternetAddress> iaList = new ArrayList<InternetAddress>(replyTo.getCount());

            for (EmailAddress emailAddress : replyTo) {
                iaList.add(new InternetAddress(emailAddress.getAddress()));
            }

            outInfo.setTargetAddresses((InternetAddress[]) iaList.toArray());
        } else if (message.getFrom() != null) {
            outInfo.setTargetAddresses(new InternetAddress[]{new InternetAddress(message.getFrom().getAddress())});
        } else {
            // does the service specify a default reply address ?
            InternetAddress replyAddress = entry.getReplyAddress();
            if (replyAddress != null) {
                outInfo.setTargetAddresses(new InternetAddress[]{replyAddress});
            }
        }

        // TODO: Add the CC Recipients
//        // save CC addresses
//        if (message.getRecipients(Message.RecipientType.CC) != null) {
//            outInfo.setCcAddresses(
//                (InternetAddress[]) message.getRecipients(Message.RecipientType.CC));
//        }

        // determine and subject for the reply message
        if (message.getSubject() != null) {
            outInfo.setSubject("Re: " + message.getSubject());
        }

        // save original message ID if one exists, so that replies can be correlated
        outInfo.setRequestMessageID(message.getInternetMessageId());
        return outInfo;
    }

    /**
     * Take specified action to either move or delete the processed email
     *
     * @param entry   the PollTableEntry for the email that has been processed
     * @param message the email message to be moved or deleted
     */
    private void moveOrDeleteAfterProcessing(final PollTableEntry entry, EwsMailClient client, EmailMessage message) throws Exception {

        String moveToFolder = null;

        switch (entry.getLastPollState()) {
            case PollTableEntry.SUCCSESSFUL:
                if (entry.getActionAfterProcess() == PollTableEntry.ActionType.MOVE) {
                    moveToFolder = entry.getMoveAfterProcess();
                }
                break;

            case PollTableEntry.FAILED:
                if (entry.getActionAfterFailure() == PollTableEntry.ActionType.MOVE) {
                    moveToFolder = entry.getMoveAfterFailure();
                }
                break;
            case PollTableEntry.NONE:
                return;
        }

        // We dont support MOVING at this moment....
        if (entry.getActionAfterProcess() == PollTableEntry.ActionType.MOVE) {
            log.error("ActionAfterProcess MOVE is currently not supported !!! Message will not be touched in mailbox");
            client.moveMessage(message, moveToFolder);
        } else if (entry.getActionAfterProcess() == PollTableEntry.ActionType.DELETE) {
            log.error("ActionAfterProcess DELETE is currently not supported !!! Message will not be touched in mailbox!");
            //client.deleteMessage(message);
        }

    }

    @Override
    protected PollTableEntry createEndpoint() {
        return new PollTableEntry(log);
    }

    public void addErrorListener(TransportErrorListener listener) {
        tess.addErrorListener(listener);
    }

    public void removeErrorListener(TransportErrorListener listener) {
        tess.removeErrorListener(listener);
    }

    /**
     * Return the UID of a message from the given folder
     *
     * @param message the message
     * @return UID as a String (long is converted to a String) or null
     */
    private String getMessageUID(EmailMessage message) throws ServiceLocalException {
        return message.getInternetMessageId();
    }
}


/* Moveing an item to another folder
Item item =new EmailMessage(service);
item.setSubject("testing move item to another folder");
item.setBody(MessageBody.getMessageBodyFromText("Item moved"));
item.setSensitivity(Sensitivity.Confidential);
item.save(new FolderId(WellKnownFolderName.Drafts));
Item item1 = Item.bind(service, item.getId());
item1.move(new FolderId(WellKnownFolderName.Notes));

 */