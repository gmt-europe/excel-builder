package nl.gmt.excelBuilder;

public class Arguments {
    private String dateFormat;
    private String sheetName;
    private String font;
    private int size;
    private String input;
    private String output;

    public Arguments(String[] args) throws ArgumentsException {
        if (args.length == 0) {
            printHelp();
        }

        boolean hadSheetName = false;
        boolean hadFont = false;
        boolean hadSize = false;
        boolean hadInput = false;
        boolean hadOutput = false;
        boolean hadDateFormat = false;

        for (String arg : args) {
            if (hadSheetName) {
                sheetName = arg;
                hadSheetName = false;
            } else if (hadFont) {
                font = arg;
                hadFont = false;
            } else if (hadSize) {
                size = Integer.parseInt(arg);
                hadSize = false;
            } else if (hadInput) {
                input = arg;
                hadInput = false;
            } else if (hadOutput) {
                output = arg;
                hadOutput = false;
            } else if (hadDateFormat) {
                dateFormat = arg;
                hadDateFormat = false;
            } else if ("-i".equals(arg)) {
                hadInput = true;
            } else if ("-o".equals(arg)) {
                hadOutput = true;
            } else if ("-s".equals(arg)) {
                hadSheetName = true;
            } else if ("-f".equals(arg)) {
                hadFont = true;
            } else if ("-S".equals(arg)) {
                hadSize = true;
            } else if ("-d".equals(arg)) {
                hadDateFormat = true;
            } else {
                throw new ArgumentsException(String.format("Invalid argument '%s'", arg));
            }
        }

        if (hadSheetName) {
            throw new ArgumentsException("Missing value for -s");
        }
        if (hadFont) {
            throw new ArgumentsException("Missing value for -f");
        }
        if (hadSize) {
            throw new ArgumentsException("Missing value for -S");
        }
        if (hadInput) {
            throw new ArgumentsException("Missing value for -i");
        }
        if (hadOutput) {
            throw new ArgumentsException("Missing value for -o");
        }
        if (hadDateFormat) {
            throw new ArgumentsException("Missing value for -d");
        }
    }

    private void printHelp() {
        System.err.println("Build an Excel from from JSON input");
        System.err.println("GMT (c) 2015");
        System.err.println();
        System.err.println("excel-builder.jar [-s <sheet name>] [-f <font>] [-S <font size>] [-i <input file name>]");
        System.err.println("                  [-o <output file name>] [-d <date format>]");
        System.err.println();
        System.err.println("  -s <sheet name>: The name of the workbook sheet; Sheet1 is used when omitted");
        System.err.println("  -f <font>: The default font; Calibri is used when omitted");
        System.err.println("  -S <font size>: The default font size; 11pt is used when omitted");
        System.err.println("  -i <input file name>: Name of the file to use as input; stdin is used");
        System.err.println("                        when omitted");
        System.err.println("  -o <output file name>: Name of the file to output the Excel file to; stdout");
        System.err.println("                         is used when omitted");
        System.err.println("  -d <date format>: The date format used to format dates; m/d/yyyy h:mm is used when");
        System.err.println("                    omitted");
        System.err.println();
    }

    public String getInput() {
        return input;
    }

    public String getOutput() {
        return output;
    }

    public String getFont() {
        return font;
    }

    public int getSize() {
        return size;
    }

    public String getSheetName() {
        return sheetName;
    }

    public String getDateFormat() {
        return dateFormat;
    }
}
