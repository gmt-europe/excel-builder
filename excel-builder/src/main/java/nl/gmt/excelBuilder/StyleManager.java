package nl.gmt.excelBuilder;

import org.apache.commons.lang.Validate;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.util.HashMap;
import java.util.Map;

public class StyleManager {
    private static final int FLAG_BOLD = 1;
    private static final int FLAG_ITALIC = 2;
    private static final int FLAG_STRIKEOUT = 4;

    private final HSSFWorkbook workbook;
    private final Map<StyleKey, HSSFCellStyle> styles = new HashMap<>();
    private final Map<FontKey, HSSFFont> fonts = new HashMap<>();
    private final String defaultFont;
    private final int defaultSize;

    public StyleManager(HSSFWorkbook workbook, String defaultFont, int defaultSize) {
        Validate.notNull(workbook, "workbook");
        Validate.notNull(defaultFont, "defaultFont");

        this.workbook = workbook;
        this.defaultFont = defaultFont;
        this.defaultSize = defaultSize;
    }

    public Builder newBuilder() {
        return new Builder();
    }

    private HSSFCellStyle buildStyle(String fontName, short size, int flags, StyleUnderline underline, StyleAlignment alignment, StyleVerticalAlignment verticalAlignment, StyleFill fill, StyleBorder border, StyleColor fontColor, StyleColor color, short dataFormat) {
        if (fontName == null) {
            fontName = defaultFont;
        }
        if (size == 0) {
            size = (short)defaultSize;
        }
        if (underline == null) {
            underline = StyleUnderline.NONE;
        }
        if (alignment == null) {
            alignment = StyleAlignment.DEFAULT;
        }
        if (verticalAlignment == null) {
            verticalAlignment = StyleVerticalAlignment.TOP;
        }
        if (fill == null) {
            fill = StyleFill.NONE;
        }
        if (border == null) {
            border = StyleBorder.NONE;
        }

        HSSFFont font = getOrCreateFont(fontName, size, flags, underline, fontColor);

        return getOrCreateStyle(alignment, verticalAlignment, fill, border, color, font, dataFormat);
    }

    private HSSFFont getOrCreateFont(String fontName, short size, int flags, StyleUnderline underline, StyleColor fontColor) {
        FontKey fontKey = new FontKey(fontName, size, flags, underline, fontColor);
        HSSFFont font = fonts.get(fontKey);
        if (font != null) {
            return font;
        }

        font = workbook.createFont();
        font.setFontName(fontName);
        font.setFontHeightInPoints(size);
        if ((flags & FLAG_BOLD) != 0) {
            font.setBold(true);
        }
        if ((flags & FLAG_ITALIC) != 0) {
            font.setItalic(true);
        }
        if ((flags & FLAG_STRIKEOUT) != 0) {
            font.setStrikeout(true);
        }
        if (underline != StyleUnderline.NONE) {
            font.setUnderline(underline.value());
        }
        if (fontColor != null) {
            font.setColor(fontColor.value());
        }

        fonts.put(fontKey, font);

        return font;
    }

    private HSSFCellStyle getOrCreateStyle(StyleAlignment alignment, StyleVerticalAlignment verticalAlignment, StyleFill fill, StyleBorder border, StyleColor color, HSSFFont font, short dataFormat) {
        StyleKey styleKey = new StyleKey(font, alignment, verticalAlignment, fill, border, color, dataFormat);
        HSSFCellStyle style = styles.get(styleKey);
        if (style != null) {
            return style;
        }

        style = workbook.createCellStyle();
        style.setFont(font);
        if (alignment != StyleAlignment.DEFAULT) {
            style.setAlignment(alignment.value());
        }
        if (verticalAlignment != StyleVerticalAlignment.TOP) {
            style.setVerticalAlignment(verticalAlignment.value());
        }
        if (fill != StyleFill.NONE) {
            style.setFillPattern(fill.value());
        }
        if (border != StyleBorder.NONE) {
            style.setBorderLeft(border.value());
            style.setBorderTop(border.value());
            style.setBorderRight(border.value());
            style.setBorderBottom(border.value());
        }
        if (color != null) {
            style.setFillForegroundColor(color.value());
        }
        if (dataFormat != 0) {
            style.setDataFormat(dataFormat);
        }

        styles.put(styleKey, style);

        return style;
    }

    private static class FontKey {
        final String font;
        final short fontSize;
        final int flags;
        private final StyleUnderline underline;
        private final StyleColor fontColor;

        public FontKey(String font, short fontSize, int flags, StyleUnderline underline, StyleColor fontColor) {
            this.font = font;
            this.fontSize = fontSize;
            this.flags = flags;
            this.underline = underline;
            this.fontColor = fontColor;
        }

        @Override
        public boolean equals(Object obj) {
            if (this == obj) {
                return true;
            }
            if (!(obj instanceof FontKey)) {
                return false;
            }

            FontKey other = (FontKey)obj;

            return fontSize == other.fontSize &&
                flags == other.flags &&
                font.equals(other.font) &&
                underline == other.underline &&
                fontColor == other.fontColor;

        }

        @Override
        public int hashCode() {
            int result = font.hashCode();
            result = 31 * result + (int)fontSize;
            result = 31 * result + flags;
            result = 31 * result + underline.hashCode();
            if (fontColor != null) {
                result = 31 * result + fontColor.hashCode();
            }
            return result;
        }
    }

    private static class StyleKey {
        private final HSSFFont font;
        final StyleAlignment alignment;
        final StyleVerticalAlignment verticalAlignment;
        final StyleFill fill;
        final StyleBorder border;
        final StyleColor color;
        private final short dataFormat;

        public StyleKey(HSSFFont font, StyleAlignment alignment, StyleVerticalAlignment verticalAlignment, StyleFill fill, StyleBorder border, StyleColor color, short dataFormat) {
            this.font = font;
            this.alignment = alignment;
            this.verticalAlignment = verticalAlignment;
            this.fill = fill;
            this.border = border;
            this.color = color;
            this.dataFormat = dataFormat;
        }

        @Override
        public boolean equals(Object obj) {
            if (this == obj) {
                return true;
            }
            if (!(obj instanceof StyleKey)) {
                return false;
            }

            StyleKey other = (StyleKey)obj;

            return font == other.font &&
                alignment == other.alignment &&
                verticalAlignment == other.verticalAlignment &&
                fill == other.fill &&
                border == other.border &&
                color == other.color &&
                dataFormat == other.dataFormat;

        }

        @Override
        public int hashCode() {
            int result = font.hashCode();
            result = 31 * result + font.hashCode();
            result = 31 * result + verticalAlignment.hashCode();
            result = 31 * result + fill.hashCode();
            result = 31 * result + border.hashCode();
            if (color != null) {
                result = 31 * result + color.hashCode();
            }
            result = 31 * result + dataFormat;
            return result;
        }
    }

    public class Builder {
        private String font;
        private short fontSize;
        private int flags;
        private StyleUnderline underline;
        private StyleAlignment alignment;
        private StyleVerticalAlignment verticalAlignment;
        private StyleFill fill;
        private StyleBorder border;
        private StyleColor fontColor;
        private StyleColor color;
        private short dataFormat;

        private Builder() {
        }

        private void setFlag(int flag, boolean set) {
            if (set) {
                this.flags |= flag;
            } else {
                this.flags &= ~flag;
            }
        }

        public Builder setFont(String font) {
            this.font = font;
            return this;
        }

        public Builder setFontSize(short fontSize) {
            this.fontSize = fontSize;
            return this;
        }

        public Builder setBold(boolean bold) {
            setFlag(FLAG_BOLD, bold);
            return this;
        }

        public Builder setItalic(boolean italic) {
            setFlag(FLAG_ITALIC, italic);
            return this;
        }

        public Builder setUnderline(StyleUnderline underline) {
            this.underline = underline;
            return this;
        }

        public Builder setStrikeout(boolean strikeout) {
            setFlag(FLAG_STRIKEOUT, strikeout);
            return this;
        }

        public Builder setAlignment(StyleAlignment alignment) {
            this.alignment = alignment;
            return this;
        }

        public Builder setVerticalAlignment(StyleVerticalAlignment verticalAlignment) {
            this.verticalAlignment = verticalAlignment;
            return this;
        }

        public Builder setFill(StyleFill fill) {
            this.fill = fill;
            return this;
        }

        public Builder setBorder(StyleBorder border) {
            this.border = border;
            return this;
        }

        public Builder setFontColor(StyleColor fontColor) {
            this.fontColor = fontColor;
            return this;
        }

        public Builder setColor(StyleColor color) {
            this.color = color;
            return this;
        }

        public Builder setDataFormat(short dataFormat) {
            this.dataFormat = dataFormat;
            return this;
        }

        public HSSFCellStyle build() {
            return buildStyle(font, fontSize, flags, underline, alignment, verticalAlignment, fill, border, fontColor, color, dataFormat);
        }
    }
}
