package convert;

/**
 * User: aturtsevitch
 * Date: May 20, 2010
 * Time: 2:42:36 PM
 */
public class InvalidDateException extends Exception
{
    public InvalidDateException()
    {
    }

    public InvalidDateException(String message)
    {
        super(message);
    }

    public InvalidDateException(String message, Throwable cause)
    {
        super(message, cause);
    }

    public InvalidDateException(Throwable cause)
    {
        super(cause);
    }
}
