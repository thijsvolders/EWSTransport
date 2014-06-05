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

package org.apache.axis2.transport.msews;

import microsoft.exchange.webservices.data.BasePropertySet;
import microsoft.exchange.webservices.data.EmailMessageSchema;
import microsoft.exchange.webservices.data.PropertySet;
import microsoft.exchange.webservices.data.WellKnownFolderName;
import org.apache.axis2.AxisFault;
import org.apache.axis2.addressing.EndpointReference;
import org.apache.axis2.description.AxisService;
import org.apache.axis2.description.Parameter;
import org.apache.axis2.description.ParameterInclude;
import org.apache.axis2.transport.base.AbstractPollTableEntry;
import org.apache.axis2.transport.base.BaseConstants;
import org.apache.axis2.transport.base.ParamUtils;
import org.apache.commons.logging.Log;

import javax.mail.internet.AddressException;
import javax.mail.internet.InternetAddress;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashSet;
import java.util.List;
import java.util.Properties;
import java.util.Set;
import java.util.StringTokenizer;

/**
 * Holds information about an entry in the VFS transport poll table used by the
 * VFS Transport Listener
 */
public class PollTableEntry extends AbstractPollTableEntry {
    private final Log log;

    // operation after mail check
    public static final int DELETE = 0;
    public static final String DELETE_VALUE = "DELETE";
    public static final int MOVE = 1;
    public static final String MOVE_VALUE = "MOVE";

    /**
     * account emailAddress to check mail
     */
    private InternetAddress emailAddress = null;
    /**
     * account password to check mail
     */
    private String password = null;
    private String serviceUrl;

    /**
     * The mail folder from which to check mail
     */
    private WellKnownFolderName folder = WellKnownFolderName.Inbox;

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
    private int actionAfterProcess = DELETE;
    /**
     * action to take after a failed poll
     */
    private int actionAfterFailure = DELETE;

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

    /**
     * UIDs of messages currently being processed
     */
    private Set<String> uidList = Collections.synchronizedSet(new HashSet<String>());

    private PropertySet ewsProperties = new PropertySet(BasePropertySet.FirstClassProperties, EmailMessageSchema.Attachments);

    public PollTableEntry(Log log) {
        this.log = log;
    }

    @Override
    public EndpointReference[] getEndpointReferences(AxisService service, String ip) {
        return new EndpointReference[]{new EndpointReference(MailConstants.TRANSPORT_PREFIX + emailAddress)};
    }

    private void addPreserveHeaders(String headerList) {
        if (headerList == null) return;
        StringTokenizer st = new StringTokenizer(headerList, " ,");
        preserveHeaders = new ArrayList<String>();
        while (st.hasMoreTokens()) {
            String token = st.nextToken();
            if (token.length() != 0) {
                preserveHeaders.add(token);
            }
        }
    }

    private void addRemoveHeaders(String headerList) {
        if (headerList == null) return;
        StringTokenizer st = new StringTokenizer(headerList, " ,");
        removeHeaders = new ArrayList<String>();
        while (st.hasMoreTokens()) {
            String token = st.nextToken();
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
        String address = ParamUtils.getOptionalParam(paramIncl, MailConstants.MAIL_EWS_EMAILADDRESS);
        if (address == null) {
            log.error("Missing " + MailConstants.MAIL_EWS_EMAILADDRESS+ " parameter in service configuration");
            return false;
        } else {
            List<Parameter> params = paramIncl.getParameters();
            Properties props = new Properties();
            for (Parameter p : params) {
                if (p.getName().startsWith("mail.")) {
                    props.setProperty(p.getName(), (String) p.getValue());
                }

                if (MailConstants.MAIL_EWS_EMAILADDRESS.equals(p.getName())) {

                    try {
                        emailAddress = new InternetAddress(address);
                    } catch (AddressException e) {
                        throw new AxisFault("Invalid email address specified by '" +
                                MailConstants.TRANSPORT_MAIL_ADDRESS + "' parameter :: " + e.getMessage());
                    }
                }

                if (MailConstants.MAIL_EWS_PASSWORD.equals(p.getName())) {
                    password = (String) p.getValue();
                }
            }

            try {
                String replyAddress = ParamUtils.getOptionalParam(paramIncl, MailConstants.TRANSPORT_MAIL_REPLY_ADDRESS);
                if (replyAddress != null) {
                    this.replyAddress = new InternetAddress(replyAddress);
                }
            } catch (AddressException e) {
                throw new AxisFault("Invalid email address specified by '" + MailConstants.TRANSPORT_MAIL_REPLY_ADDRESS + "' parameter :: " + e.getMessage());
            }

            String transportFolderNameValue = ParamUtils.getOptionalParam(paramIncl, MailConstants.TRANSPORT_MAIL_FOLDER);
            folder = WellKnownFolderName.valueOf(transportFolderNameValue == null ? folder.name() : transportFolderNameValue);

            addPreserveHeaders(ParamUtils.getOptionalParam(paramIncl, MailConstants.TRANSPORT_MAIL_PRESERVE_HEADERS));
            addRemoveHeaders(ParamUtils.getOptionalParam(paramIncl, MailConstants.TRANSPORT_MAIL_REMOVE_HEADERS));

            String option = ParamUtils.getOptionalParam(paramIncl, MailConstants.TRANSPORT_MAIL_ACTION_AFTER_PROCESS);
            actionAfterProcess = PollTableEntry.MOVE_VALUE.equalsIgnoreCase(option) ? PollTableEntry.MOVE : PollTableEntry.DELETE;

            option = ParamUtils.getOptionalParam(paramIncl, MailConstants.TRANSPORT_MAIL_ACTION_AFTER_FAILURE);
            actionAfterFailure = PollTableEntry.MOVE_VALUE.equalsIgnoreCase(option) ? PollTableEntry.MOVE : PollTableEntry.DELETE;

            moveAfterProcess = ParamUtils.getOptionalParam(paramIncl, MailConstants.TRANSPORT_MAIL_MOVE_AFTER_PROCESS);
            moveAfterFailure = ParamUtils.getOptionalParam(paramIncl, MailConstants.TRANSPORT_MAIL_MOVE_AFTER_FAILURE);

            String processInParallel = ParamUtils.getOptionalParam(paramIncl, MailConstants.TRANSPORT_MAIL_PROCESS_IN_PARALLEL);
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

            // EWS configuration
            password = ParamUtils.getRequiredParam(paramIncl, MailConstants.MAIL_EWS_PASSWORD);
            serviceUrl = ParamUtils.getRequiredParam(paramIncl, MailConstants.MAIL_EWS_URL);

            String msgCountParamValue = ParamUtils.getOptionalParam(paramIncl, MailConstants.MAIL_EWS_MAX_MSG_COUNT);
            // When msgCountParamValue not an integer then an exception will be thrown. Thats good! :)
            messageCount = msgCountParamValue == null ? messageCount : Integer.parseInt(msgCountParamValue);

            return super.loadConfiguration(paramIncl);
        }
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

    public WellKnownFolderName getFolder() {
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

    public int getActionAfterProcess() {
        return actionAfterProcess;
    }

    public int getActionAfterFailure() {
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

    public int getMessageCount() {
        return messageCount;
    }

    public String getAttachmentRegExp() {
        return attachmentRegExp;
    }

    public PropertySet getEwsProperties() {
        return ewsProperties;
    }
}
