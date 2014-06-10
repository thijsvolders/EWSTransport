package nl.yenlo.transport.msews.log;

import microsoft.exchange.webservices.data.ITraceListener;
import org.apache.commons.logging.Log;

/**
 * @author tvolders.
 */
public class EWSTraceListener implements ITraceListener {
    private Log log;

    public EWSTraceListener(Log log) {
        this.log = log;
    }

    public void trace(String traceType, String traceMessage) {
        log.trace(traceMessage);
    }
}
