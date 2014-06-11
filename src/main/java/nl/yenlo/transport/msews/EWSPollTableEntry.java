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

import microsoft.exchange.webservices.data.FolderId;
import microsoft.exchange.webservices.data.WellKnownFolderName;
import nl.yenlo.transport.msews.client.exception.EwsMailClientConfigException;
import org.apache.axis2.AxisFault;
import org.apache.axis2.addressing.EndpointReference;
import org.apache.axis2.description.AxisService;
import org.apache.axis2.description.ParameterInclude;
import org.apache.axis2.description.TransportInDescription;
import org.apache.axis2.transport.base.AbstractPollTableEntry;
import org.apache.axis2.transport.base.BaseConstants;
import org.apache.axis2.transport.base.ParamUtils;
import org.apache.commons.lang.StringUtils;
import org.apache.commons.logging.Log;

import javax.mail.internet.AddressException;
import javax.mail.internet.InternetAddress;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.StringTokenizer;

/**
 * Holds information about an entry in the VFS transport poll table used by the
 * VFS Transport Listener
 */
public class EWSPollTableEntry extends AbstractPollTableEntry {
    private final Log log;

    /**
     * The supported ExtractTypes.
     */
    public enum ExtractType {
        /**
         * Extract the mails' body
         */
        BODY,

        /**
         * Extract the mail's (one of) attachment.
         */
        ATTACHMENTS
    }


    /**
     * The supported actions which are taken when a mail has been processed (successfully or failed)
     */
    public enum ActionType {
        /**
         * Move the mail to another mailbox folder
         */
        MOVE,
        /**
         * Delete the mail from the current mailbox folder
         */
        DELETE,

        /**
         * Mark the mail as read
         */
        MARKASREAD,

        /**
         * Do nothing. leave the mail be.
         */
        NOTHING
    }


    /**
     * When we delete something will it be forced (really deleted) or moved to trash !?
     */
    public enum DeleteActionType {
        /**
         * FORCE means non-recoverable
         */
        FORCE,

        /**
         * Move mail to trash when deleting
         */
        TRASH
    }


    /**
     * account emailAddress to check mail
     */
    private InternetAddress emailAddress = null;

    /**
     * account password to check mail
     */
    private String password = null;

    /**
     * domain of the provided account.
     */
    private String domain = null;

    /**
     * The service url of the EWS service.
     * i.e.: https://ExchangeMailHost/EWS/Exchange.asmx
     */
    private String serviceUrl;

    /**
     * The mail folder from which to check mail
     */
    private FolderId folder = new FolderId(WellKnownFolderName.Inbox);

    /**
     * default reply address
     */
    private InternetAddress replyAddress = null;

    /**
     * list of mail headers to be preserved into the Axis2 message as transport headers
     */
    private List<String> preserveHeaders = null;
    /**
     * list of mail headers to be removed from the Axis2 message transport headers
     */
    private List<String> removeHeaders = null;

    /**
     * action to take after a successful poll
     */
    private ActionType actionAfterProcess = ActionType.MARKASREAD;
    /**
     * action to take after a failed poll
     */
    private ActionType actionAfterFailure = ActionType.NOTHING;

    private DeleteActionType deleteActionType = DeleteActionType.TRASH; // Use TrashDelete per default

    /**
     * folder to move the email after processing
     */
    private String moveAfterProcess;
    /**
     * folder to move the email after failure
     */
    private String moveAfterFailure;
    /**
     * Should mail be processed in parallel? e.g. with IMAP
     */
    private boolean processingMailInParallel = false;

    /**
     * Process 10 mails per mail server poll.
     */
    private int messageCount = 10;

    // FIXME: Add Attachment Selection RegExp pattern (filename regexp)
    private String attachmentRegExp = null;

    private ExtractType extractType = ExtractType.BODY;


    /**
     * UIDs of messages currently being processed
     */
    private Set<String> uidList = Collections.synchronizedSet(new HashSet<String>());

    public EWSPollTableEntry(Log log) {
        this.log = log;
    }

    @Override
    public EndpointReference[] getEndpointReferences(AxisService service, String ip) {
        return new EndpointReference[]{new EndpointReference("ews" + emailAddress)};
    }

    private void addPreserveHeaders(String headerList) {
        if (headerList == null) return;
        StringTokenizer st = new StringTokenizer(headerList, ",");
        preserveHeaders = new ArrayList<String>();
        while (st.hasMoreTokens()) {
            String token = st.nextToken().trim();
            if (token.length() != 0) {
                preserveHeaders.add(token);
            }
        }
    }

    private void addRemoveHeaders(String headerList) {
        if (headerList == null) return;
        StringTokenizer st = new StringTokenizer(headerList, ",");
        removeHeaders = new ArrayList<String>();
        while (st.hasMoreTokens()) {
            String token = st.nextToken().trim();
            if (token.length() != 0) {
                removeHeaders.add(token);
            }
        }
    }

    public boolean retainHeader(String name) {
        if (preserveHeaders != null) {
            return preserveHeaders.contains(name);
        } else if (removeHeaders != null) {
            return !removeHeaders.contains(name);
        } else {
            return true;
        }
    }

    @Override
    public boolean loadConfiguration(ParameterInclude paramIncl) throws AxisFault {

        if (paramIncl instanceof TransportInDescription) {
            // This is called when the transport is first initialized (at server start)...
            // We dont initialize the transport at this stage...
            return false;
        } else {

            String address = ParamUtils.getRequiredParam(paramIncl, EWSTransportConstants.MAIL_EWS_EMAILADDRESS);
            try {
                emailAddress = new InternetAddress(address);
            } catch (AddressException e) {
                throw new AxisFault("Invalid email address specified by '" + EWSTransportConstants.MAIL_EWS_EMAILADDRESS + "' parameter :: " + e.getMessage());
            }

            password = ParamUtils.getRequiredParam(paramIncl, EWSTransportConstants.MAIL_EWS_PASSWORD);

            try {
                String replyAddress = ParamUtils.getOptionalParam(paramIncl, EWSTransportConstants.TRANSPORT_MAIL_REPLY_ADDRESS);
                if (replyAddress != null) {
                    this.replyAddress = new InternetAddress(replyAddress);
                }
            } catch (AddressException e) {
                throw new AxisFault("Invalid email address specified by '" + EWSTransportConstants.TRANSPORT_MAIL_REPLY_ADDRESS + "' parameter :: " + e.getMessage());
            }

            String transportFolderNameValue = ParamUtils.getOptionalParam(paramIncl, EWSTransportConstants.MAIL_EWS_FOLDER);
            try {
                if (transportFolderNameValue != null) {
                    // Test if the supplied transportFolderName is a WellknownFolderName, use it if so..
                    try {
                        folder = new FolderId(WellKnownFolderName.valueOf(transportFolderNameValue));
                    } catch (EnumConstantNotPresentException ecnpe) {
                        // OK, no known name.. FolderId must be UniqueId
                        folder = new FolderId(transportFolderNameValue);
                    }
                }
            } catch (Exception e) {
                throw new EwsMailClientConfigException("The " + EWSTransportConstants.MAIL_EWS_FOLDER + " parameters is either not found or null", e);
            }

            // EWS configuration
            password = ParamUtils.getRequiredParam(paramIncl, EWSTransportConstants.MAIL_EWS_PASSWORD);
            serviceUrl = ParamUtils.getRequiredParam(paramIncl, EWSTransportConstants.MAIL_EWS_URL);
            domain = ParamUtils.getRequiredParam(paramIncl, EWSTransportConstants.MAIL_EWS_DOMAIN);

            addPreserveHeaders(ParamUtils.getOptionalParam(paramIncl, EWSTransportConstants.TRANSPORT_MAIL_PRESERVE_HEADERS));
            addRemoveHeaders(ParamUtils.getOptionalParam(paramIncl, EWSTransportConstants.TRANSPORT_MAIL_REMOVE_HEADERS));

            try {
                String option = ParamUtils.getOptionalParam(paramIncl, EWSTransportConstants.TRANSPORT_MAIL_ACTION_AFTER_PROCESS);
                if (option != null) {
                    actionAfterProcess = ActionType.valueOf(option);
                }
            } catch (EnumConstantNotPresentException ecnpe) {
                log.error("The supplied " + EWSTransportConstants.TRANSPORT_MAIL_ACTION_AFTER_PROCESS + " is not supported. Please use one of the following " + StringUtils.join(ActionType.values()));
                throw ecnpe;
            }

            try {
                String option = ParamUtils.getOptionalParam(paramIncl, EWSTransportConstants.TRANSPORT_MAIL_ACTION_AFTER_FAILURE);
                if (option != null) {
                    actionAfterFailure = ActionType.valueOf(option);
                }
            } catch (EnumConstantNotPresentException ecnpe) {
                log.error("The supplied " + EWSTransportConstants.TRANSPORT_MAIL_ACTION_AFTER_FAILURE + " is not supported. Please use one of the following " + StringUtils.join(ActionType.values()));
                throw ecnpe;
            }

            moveAfterProcess = ParamUtils.getOptionalParam(paramIncl, EWSTransportConstants.TRANSPORT_MAIL_MOVE_AFTER_PROCESS);
            moveAfterFailure = ParamUtils.getOptionalParam(paramIncl, EWSTransportConstants.TRANSPORT_MAIL_MOVE_AFTER_FAILURE);

            String processInParallel = ParamUtils.getOptionalParam(paramIncl, EWSTransportConstants.TRANSPORT_MAIL_PROCESS_IN_PARALLEL);
            if (processInParallel != null) {
                processingMailInParallel = Boolean.parseBoolean(processInParallel);
                if (log.isDebugEnabled() && processingMailInParallel) {
                    log.debug("Parallel mail processing enabled for : " + address);
                }
            }

            String pollInParallel = ParamUtils.getOptionalParam(paramIncl, BaseConstants.TRANSPORT_POLL_IN_PARALLEL);
            if (pollInParallel != null) {
                setConcurrentPollingAllowed(Boolean.parseBoolean(pollInParallel));
                if (log.isDebugEnabled() && isConcurrentPollingAllowed()) {
                    log.debug("Concurrent mail polling enabled for : " + address);
                }
            }

            String msgCountParamValue = ParamUtils.getOptionalParam(paramIncl, EWSTransportConstants.MAIL_EWS_MAX_MSG_COUNT);
            // When msgCountParamValue not an integer then an exception will be thrown. Thats good! :)
            messageCount = msgCountParamValue == null ? messageCount : Integer.parseInt(msgCountParamValue);

            String optionalParam = ParamUtils.getOptionalParam(paramIncl, EWSTransportConstants.TRANSPORT_MAIL_EXTRACTTYPE);
            try {
                if (optionalParam != null) {
                    extractType = ExtractType.valueOf(optionalParam);
                }
            } catch (EnumConstantNotPresentException ecnpe) {
                log.error("The supplied " + EWSTransportConstants.TRANSPORT_MAIL_EXTRACTTYPE + " is not supported. Please use one of the following " + StringUtils.join(ExtractType.values()));
                throw ecnpe;
            }

            optionalParam = ParamUtils.getOptionalParam(paramIncl, EWSTransportConstants.TRANSPORT_MAIL_DELETETYPE);
            try {
                if (optionalParam != null) {
                    deleteActionType = DeleteActionType.valueOf(optionalParam);
                }
            } catch (EnumConstantNotPresentException ecnpe) {
                log.error("The supplied " + EWSTransportConstants.TRANSPORT_MAIL_DELETETYPE + " is not supported. Please use one of the following " + StringUtils.join(DeleteActionType.values()));
                throw ecnpe;
            }

        }
        return super.loadConfiguration(paramIncl);
    }


    public synchronized void processingUID(String uid) {
        this.uidList.add(uid);
    }

    public synchronized boolean isProcessingUID(String uid) {
        return this.uidList.contains(uid);
    }

    public synchronized void removeUID(String uid) {
        this.uidList.remove(uid);
    }


    /*
     * Getters and setters
     */
    public InternetAddress getEmailAddress() {
        return emailAddress;
    }

    public String getPassword() {
        return password;
    }

    public String getServiceUrl() {
        return serviceUrl;
    }

    public String getDomain() {
        return domain;
    }

    public int getMessageCount() {
        return messageCount;
    }

    public FolderId getFolder() {
        return folder;
    }

    public InternetAddress getReplyAddress() {
        return replyAddress;
    }

    public List<String> getPreserveHeaders() {
        return preserveHeaders;
    }

    public List<String> getRemoveHeaders() {
        return removeHeaders;
    }

    public ActionType getActionAfterProcess() {
        return actionAfterProcess;
    }

    public ActionType getActionAfterFailure() {
        return actionAfterFailure;
    }

    public String getMoveAfterProcess() {
        return moveAfterProcess;
    }

    public String getMoveAfterFailure() {
        return moveAfterFailure;
    }

    public boolean isProcessingMailInParallel() {
        return processingMailInParallel;
    }

    public void setProcessingMailInParallel(boolean processingMailInParallel) {
        this.processingMailInParallel = processingMailInParallel;
    }

    public ExtractType getExtractType() {
        return extractType;
    }

    public DeleteActionType getDeleteActionType() {
        return deleteActionType;
    }
}
