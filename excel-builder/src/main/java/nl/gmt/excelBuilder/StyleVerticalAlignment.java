package nl.gmt.excelBuilder;

import org.apache.commons.lang.Validate;
import org.apache.poi.ss.usermodel.CellStyle;

public enum StyleVerticalAlignment {
    TOP("top", CellStyle.VERTICAL_TOP),
    CENTER("center", CellStyle.VERTICAL_CENTER),
    BOTTOM("bottom", CellStyle.VERTICAL_BOTTOM),
    JUSTIFY("justify", CellStyle.VERTICAL_JUSTIFY);

    private final String code;
    private final short value;

    StyleVerticalAlignment(String code, short value) {
        this.code = code;
        this.value = value;
    }

    public short value() {
        return value;
    }

    public static StyleVerticalAlignment findByCode(String code) {
        Validate.notNull(code, "code");

        for (StyleVerticalAlignment alignment : values()) {
            if (code.equals(alignment.code)) {
                return alignment;
            }
        }

        return null;
    }
}
