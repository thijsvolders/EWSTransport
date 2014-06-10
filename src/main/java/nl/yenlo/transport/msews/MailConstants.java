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

import microsoft.exchange.webservices.data.WellKnownFolderName;

import javax.mail.Session;

public class MailConstants {
    public static WellKnownFolderName DEFAULT_FOLDER = WellKnownFolderName.Inbox;

    public static final String TRANSPORT_MAIL_ACTION_AFTER_PROCESS = "transport.mail.ActionAfterProcess";
    public static final String TRANSPORT_MAIL_ACTION_AFTER_FAILURE = "transport.mail.ActionAfterFailure";
    public static final String TRANSPORT_MAIL_DELETE_ACTION_TYPE = "transport.mail.DeleteActionType";

    public static final String TRANSPORT_MAIL_MOVE_AFTER_PROCESS = "transport.mail.MoveAfterProcess";
    public static final String TRANSPORT_MAIL_MOVE_AFTER_FAILURE = "transport.mail.MoveAfterFailure";

    public static final String TRANSPORT_MAIL_PROCESS_IN_PARALLEL = "transport.mail.ProcessInParallel";

    public static final String TRANSPORT_MAIL_ADDRESS  = "transport.mail.Address";
    
    public static final String TRANSPORT_MAIL_DEBUG = "transport.mail.Debug";
    
    public static final String TRANSPORT_MAIL_FOLDER           = "transport.mail.Folder";
    public static final String TRANSPORT_MAIL_REPLY_ADDRESS    = "transport.mail.ReplyAddress";
    public static final String TRANSPORT_MAIL_PRESERVE_HEADERS = "transport.mail.PreserveHeaders";
    public static final String TRANSPORT_MAIL_REMOVE_HEADERS   = "transport.mail.RemoveHeaders";
    public static final String TRANSPORT_MAIL_EXTRACTTYPE = "transport.mail.ExtractType";

    // POP3 and IMAP properties
    public static final String MAIL_EWS_EMAILADDRESS = "mail.ews.email";
    public static final String MAIL_EWS_PASSWORD = "mail.ews.password";
    public static final String MAIL_EWS_URL = "mail.ews.url";
    public static final String MAIL_EWS_MAX_MSG_COUNT = "mail.ews.maxMessageCount";

    // transport / mail headers
    public static final String MAIL_HEADER_TO          = "To";
    public static final String MAIL_HEADER_FROM        = "From";
    public static final String MAIL_HEADER_CC          = "Cc";
    public static final String MAIL_HEADER_BCC         = "Bcc";
    public static final String MAIL_HEADER_REPLY_TO    = "Reply-To";
    public static final String MAIL_HEADER_IN_REPLY_TO = "In-Reply-To";
    public static final String MAIL_HEADER_SUBJECT     = "Subject";
    public static final String MAIL_HEADER_MESSAGE_ID  = "Message-ID";
    public static final String MAIL_HEADER_REFERENCES  = "References";

    // Custom headers
    public static final String TRANSPORT_MAIL_CUSTOM_HEADERS     = "transport.mail.custom.headers";
    
}
