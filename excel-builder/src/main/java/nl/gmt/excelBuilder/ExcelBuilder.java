package nl.gmt.excelBuilder;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.json.JSONObject;
import org.json.JSONTokener;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Calendar;
import java.util.Iterator;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ExcelBuilder {
    private static final Pattern DATE_PATTERN = Pattern.compile("^(\\d{4})-(\\d{1,2})-(\\d{1,2})T(\\d{1,2}):(\\d{1,2}):(\\d{1,2})(?:\\.(\\d{1,3}))?$");
    private final Arguments arguments;
    private final InputStream is;
    private final OutputStream os;

    public ExcelBuilder(Arguments arguments, InputStream is, OutputStream os) {
        this.arguments = arguments;
        this.is = is;
        this.os = os;
    }

    @SuppressWarnings("MagicConstant")
    public void build() throws IOException {
        JSONObject json = new JSONObject(new JSONTokener(is));

        HSSFWorkbook workbook = new HSSFWorkbook();

        String sheetName = arguments.getSheetName();
        if (sheetName == null) {
            sheetName = "Sheet1";
        }
        HSSFSheet sheet = workbook.createSheet(sheetName);

        String defaultFont = arguments.getFont();
        if (defaultFont == null) {
            defaultFont = "Calibri";
        }
        int defaultSize = arguments.getSize();
        if (defaultSize == 0) {
            defaultSize = 11;
        }

        String dateFormatFormat = arguments.getDateFormat();
        if (dateFormatFormat == null) {
            dateFormatFormat = "m/d/yyyy h:mm";
        }

        short dateFormat = workbook.getCreationHelper()
            .createDataFormat()
            .getFormat(dateFormatFormat);

        HSSFFont workbookFont = workbook.getFontAt((short)0);
        workbookFont.setFontName(defaultFont);
        workbookFont.setFontHeightInPoints((short)defaultSize);

        StyleManager styleManager = new StyleManager(workbook, defaultFont, defaultSize);

        for (Iterator iter = json.keys(); iter.hasNext(); ) {
            String key = (String)iter.next();
            Object value = json.get(key);
            JSONObject values = null;
            if (value instanceof JSONObject) {
                values = (JSONObject)value;
                value = null;
            }

            Index index = Index.parse(key);
            if (index == null) {
                System.out.println("Cannot parse cell '" + key + "'");
                continue;
            }

            HSSFRow row = sheet.getRow(index.row);
            if (row == null) {
                row = sheet.createRow(index.row);
            }
            HSSFCell cell = row.getCell(index.column);
            if (cell == null) {
                cell = row.createCell(index.column);
            }

            StyleManager.Builder style = styleManager.newBuilder();
            int columnSpan = -1;
            int rowSpan = -1;
            int height = -1;
            int width = -1;

            if (values != null) {
                for (Iterator valuesIter = values.keys(); valuesIter.hasNext(); ) {
                    String property = (String)valuesIter.next();

                    switch (property) {
                        case "value":
                            value = values.get(property);
                            break;

                        case "font":
                            style.setFont(values.getString(property));
                            break;

                        case "size":
                        case "fontSize":
                            style.setFontSize((short)values.getInt(property));
                            break;

                        case "bold":
                            style.setBold(values.getBoolean(property));
                            break;

                        case "italic":
                            style.setItalic(values.getBoolean(property));
                            break;

                        case "strikeout":
                            style.setStrikeout(values.getBoolean(property));
                            break;

                        case "underline":
                            StyleUnderline underline = StyleUnderline.findByCode(values.getString(property));
                            if (underline == null) {
                                System.err.println("Cannot parse underline '" + values.getString(property) + "'");
                            } else {
                                style.setUnderline(underline);
                            }
                            break;

                        case "align":
                        case "alignment":
                            StyleAlignment alignment = StyleAlignment.findByCode(values.getString(property));
                            if (alignment == null) {
                                System.err.println("Cannot parse alignment '" + values.getString(property) + "'");
                            } else {
                                style.setAlignment(alignment);
                            }
                            break;

                        case "valign":
                        case "verticalAlignment":
                            StyleVerticalAlignment verticalAlignment = StyleVerticalAlignment.findByCode(values.getString(property));
                            if (verticalAlignment == null) {
                                System.err.println("Cannot parse vertical alignment '" + values.getString(property) + "'");
                            } else {
                                style.setVerticalAlignment(verticalAlignment);
                            }
                            break;

                        case "fill":
                            StyleFill fill = StyleFill.findByCode(values.getString(property));
                            if (fill == null) {
                                System.err.println("Cannot parse fill '" + values.getString(property) + "'");
                            } else {
                                style.setFill(fill);
                            }
                            break;

                        case "border":
                            StyleBorder border = StyleBorder.findByCode(values.getString(property));
                            if (border == null) {
                                System.err.println("Cannot parse border '" + values.getString(property) + "'");
                            } else {
                                style.setBorder(border);
                            }
                            break;

                        case "fontColor":
                        case "foreColor":
                        case "textColor":
                            StyleColor fontColor = StyleColor.findByCode(values.getString(property));
                            if (fontColor == null) {
                                System.err.println("Cannot parse fontColor '" + values.getString(property) + "'");
                            } else {
                                style.setFontColor(fontColor);
                            }
                            break;

                        case "color":
                        case "backColor":
                            StyleColor color = StyleColor.findByCode(values.getString(property));
                            if (color == null) {
                                System.err.println("Cannot parse color '" + values.getString(property) + "'");
                            } else {
                                style.setColor(color);
                            }
                            break;

                        case "height":
                            height = values.getInt(property);
                            break;

                        case "width":
                            width = values.getInt(property);
                            break;

                        case "span":
                        case "colspan":
                        case "columnSpan":
                            columnSpan = values.getInt(property);
                            break;

                        case "vspan":
                        case "rowspan":
                        case "rowSpan":
                            rowSpan = values.getInt(property);
                            break;

                        default:
                            System.err.println("Cannot process property '" + property + "'");
                    }
                }
            }

            if (value instanceof String) {
                Matcher matcher = DATE_PATTERN.matcher((String)value);
                if (matcher.matches()) {
                    Calendar calendar = Calendar.getInstance();
                    calendar.set(Calendar.YEAR, Integer.parseInt(matcher.group(1)));
                    calendar.set(Calendar.MONTH, Integer.parseInt(matcher.group(2)));
                    calendar.set(Calendar.DAY_OF_MONTH, Integer.parseInt(matcher.group(3)) - 1);
                    calendar.set(Calendar.HOUR_OF_DAY, Integer.parseInt(matcher.group(4)));
                    calendar.set(Calendar.MINUTE, Integer.parseInt(matcher.group(5)));
                    calendar.set(Calendar.SECOND, Integer.parseInt(matcher.group(6)));
                    if (matcher.group(7) != null) {
                        calendar.set(Calendar.MILLISECOND, Integer.parseInt(matcher.group(7)));
                    }
                    cell.setCellValue(calendar);

                    style.setDataFormat(dateFormat);
                } else {
                    cell.setCellValue((String)value);
                }
            } else if (value instanceof Number) {
                cell.setCellValue(((Number)value).doubleValue());
            } else if (value instanceof Boolean) {
                cell.setCellValue((Boolean)value);
            } else if (value != null) {
                System.err.println("Cannot set cell value to value of type '" + value.getClass().getName() + "'");
            }

            cell.setCellStyle(style.build());

            if (height != -1) {
                row.setHeight((short)Math.max(row.getHeight(), height));
            }
            if (width != -1) {
                sheet.setColumnWidth(index.column, Math.max(sheet.getColumnWidth(index.column), width));
            }
            if (columnSpan > 1 || rowSpan > 1) {
                sheet.addMergedRegion(new CellRangeAddress(
                    index.row,
                    index.row + Math.max(0, rowSpan - 1),
                    index.column,
                    index.column + Math.max(0, columnSpan - 1)
                ));
            }
        }

        workbook.write(os);
    }

    private static class Index {
        final int row;
        final int column;

        public Index(int row, int column) {
            this.row = row;
            this.column = column;
        }

        static Index parse(String cell) {
            int column = -1;
            int row = -1;
            boolean afterColumn = false;

            for (int i = 0; i < cell.length(); i++) {
                char c = cell.charAt(i);
                if (afterColumn) {
                    if (row == -1) {
                        row = 0;
                    }

                    if (c >= '0' && c <= '9') {
                        row *= 10;
                        row += (int)c - (int)'0';
                    } else {
                        return null;
                    }
                } else {
                    if (c >= '0' && c <= '9') {
                        if (column == -1) {
                            return null;
                        }
                        afterColumn = true;
                        i--;
                    } else {
                        if (column == -1) {
                            column = 0;
                        }
                        if (c >= 'a' && c <= 'z') {
                            column *= 26;
                            column += (int)c - (int)'a';
                        } else if (c >= 'A' && c <= 'Z') {
                            column *= 26;
                            column += (int)c - (int)'A';
                        } else {
                            return null;
                        }
                    }
                }
            }

            if (column == -1 || row == -1) {
                return null;
            }

            return new Index(row - 1, column);
        }
    }
}
