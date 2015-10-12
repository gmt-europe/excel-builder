package nl.gmt.excelBuilder;

public class ArgumentsException extends Exception {
    public ArgumentsException() {
    }

    public ArgumentsException(String s) {
        super(s);
    }

    public ArgumentsException(String s, Throwable throwable) {
        super(s, throwable);
    }

    public ArgumentsException(Throwable throwable) {
        super(throwable);
    }

    public ArgumentsException(String s, Throwable throwable, boolean b, boolean b1) {
        super(s, throwable, b, b1);
    }
}
