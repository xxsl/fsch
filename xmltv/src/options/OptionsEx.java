package options;

import org.apache.commons.cli.Option;
import org.apache.commons.cli.Options;


public class OptionsEx extends Options
{
    @Override
    public Options addOption(String opt, boolean hasArg, String description)
    {
        Option option = new OptionEx(opt, hasArg, description);
        return addOption(option);
    }

    @Override
    public Options addOption(String opt, String longOpt, boolean hasArg, String description)
    {
        Option option = new OptionEx(opt, longOpt, hasArg, description);
        return addOption(option);
    }

    public Options addOption(String opt, boolean hasArg, String description, String defaultValue)
    {
        Option option = new OptionEx(opt, hasArg, description, defaultValue);
        return addOption(option);
    }

    public Options addOption(String opt, String longOpt, boolean hasArg, String description, String defaultValue)
    {
        Option option = new OptionEx(opt, longOpt, hasArg, description, defaultValue);
        return addOption(option);
    }

    public OptionEx getOptionEx(String opt)
    {
        return (OptionEx)super.getOption(opt);    
    }
}
