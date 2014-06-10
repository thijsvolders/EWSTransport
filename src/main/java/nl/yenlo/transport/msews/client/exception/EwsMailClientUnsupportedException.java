package nl.yenlo.transport.msews.client.exception;

/**
 * @author tvolders.
 */
public class EwsMailClientUnsupportedException extends RuntimeException {
    public EwsMailClientUnsupportedException(String message, Throwable cause) {
        super(message, cause);
    }
}
