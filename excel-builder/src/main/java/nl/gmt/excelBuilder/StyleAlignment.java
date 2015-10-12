package nl.gmt.excelBuilder;

import org.apache.commons.lang.Validate;
import org.apache.poi.ss.usermodel.CellStyle;

public enum StyleAlignment {
    DEFAULT("default", CellStyle.ALIGN_GENERAL),
    LEFT("left", CellStyle.ALIGN_LEFT),
    CENTER("center", CellStyle.ALIGN_CENTER),
    RIGHT("right", CellStyle.ALIGN_RIGHT),
    JUSTIFY("justify", CellStyle.ALIGN_JUSTIFY);

    private final String code;
    private final short value;

    StyleAlignment(String code, short value) {
        this.code = code;
        this.value = value;
    }

    public short value() {
        return value;
    }

    public static StyleAlignment findByCode(String code) {
        Validate.notNull(code, "code");

        for (StyleAlignment alignment : values()) {
            if (code.equals(alignment.code)) {
                return alignment;
            }
        }

        return null;
    }
}
