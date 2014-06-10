package nl.yenlo.transport.msews.client.exception;

/**
 * @author tvolders.
 */
public class EwsMailClientCommunicationException extends RuntimeException {
    public EwsMailClientCommunicationException(String message, Throwable cause) {
        super(message, cause);
    }
}
