public class Arg {
    private String value;

    public Arg() {
    }

    public String getValue() {
        return value;
    }

    public void setValue(String value) {
        this.value = value;
    }


    @Override
    public String toString() {
        return value;
    }
}
