package options;

import org.apache.commons.cli.Option;

/**
 * User: aturtsevitch
 * Date: May 20, 2010
 * Time: 1:04:02 PM
 */
public class OptionEx extends Option
{
    public OptionEx(String opt, String description)
            throws IllegalArgumentException
    {
        super(opt, description);
    }

    public OptionEx(String opt, boolean hasArg, String description)
            throws IllegalArgumentException
    {
        super(opt, hasArg, description);
    }

    public OptionEx(String opt, String longOpt, boolean hasArg, String description)
            throws IllegalArgumentException
    {
        super(opt, longOpt, hasArg, description);
    }
}
