package nl.gmt.excelBuilder;

import org.apache.commons.lang.Validate;
import org.apache.poi.ss.usermodel.CellStyle;

public enum StyleBorder {
    NONE("none", CellStyle.BORDER_NONE),
    THIN("thin", CellStyle.BORDER_THIN),
    MEDIUM("medium", CellStyle.BORDER_MEDIUM),
    DASHED("dashed", CellStyle.BORDER_DASHED),
    HAIR("hair", CellStyle.BORDER_HAIR),
    THICK("thick", CellStyle.BORDER_THICK),
    DOUBLE("double", CellStyle.BORDER_DOUBLE),
    DOTTED("dotted", CellStyle.BORDER_DOTTED),
    MEDIUM_DASHED("medium dashed", CellStyle.BORDER_MEDIUM_DASHED),
    DASH_DOT("dash dot", CellStyle.BORDER_DASH_DOT),
    MEDIUM_DASH_DOT("medium dash dot", CellStyle.BORDER_MEDIUM_DASH_DOT),
    DASH_DOT_DOT("dash dot dot", CellStyle.BORDER_DASH_DOT_DOT),
    MEDIUM_DASH_DOT_DOT("medium dash dot dot", CellStyle.BORDER_MEDIUM_DASH_DOT_DOT),
    SLANTED_DASH_DOT("slanted dash dot", CellStyle.BORDER_SLANTED_DASH_DOT);

    private final String code;
    private final short value;

    StyleBorder(String code, short value) {
        this.code = code;
        this.value = value;
    }

    public short value() {
        return value;
    }

    public static StyleBorder findByCode(String code) {
        Validate.notNull(code, "code");

        for (StyleBorder alignment : values()) {
            if (code.equals(alignment.code)) {
                return alignment;
            }
        }

        return null;
    }
}
