package options;

import org.apache.commons.cli.Option;

public class OptionEx extends Option
{
    private String defaultValue;

    public OptionEx(String opt, String description) throws IllegalArgumentException
    {
        super(opt, description);
    }

    public OptionEx(String opt, boolean hasArg, String description) throws IllegalArgumentException
    {
        super(opt, hasArg, description);
    }

    public OptionEx(String opt, String longOpt, boolean hasArg, String description) throws IllegalArgumentException
    {
        super(opt, longOpt, hasArg, description);
    }

    public OptionEx(String opt, String description, String defaultValue) throws IllegalArgumentException
    {
        super(opt, description);
        this.defaultValue = defaultValue;
    }

    public OptionEx(String opt, boolean hasArg, String description, String defaultValue) throws IllegalArgumentException
    {
        super(opt, hasArg, description);
        this.defaultValue = defaultValue;
    }

    public OptionEx(String opt, String longOpt, boolean hasArg, String description, String defaultValue) throws IllegalArgumentException
    {
        super(opt, longOpt, hasArg, description);
        this.defaultValue = defaultValue;
    }

    public String getDefaultValue()
    {
        return defaultValue;
    }

    @Override
    public String getDescription()
    {
        return super.getDescription() + (getDefaultValue() != null ? ". Default value: " + getDefaultValue() : "");
    }
}
