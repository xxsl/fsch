package options;

import org.apache.commons.cli.Option;
import org.apache.commons.cli.Options;

/**
 * User: aturtsevitch
 * Date: May 20, 2010
 * Time: 1:03:46 PM
 */
public class OptionsEx extends Options
{
    @Override
    public Options addOption(String opt, boolean hasArg, String description)
    {
        return super.addOption(opt, hasArg, description);    //TODO : impl
    }

    @Override
    public Options addOption(String opt, String longOpt, boolean hasArg, String description)
    {
        return super.addOption(opt, longOpt, hasArg, description);    //TODO : impl
    }

    @Override
    public Options addOption(Option opt)
    {
        return super.addOption(opt);    //TODO : impl
    }
}
