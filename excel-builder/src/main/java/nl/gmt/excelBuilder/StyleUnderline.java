package nl.gmt.excelBuilder;

import org.apache.commons.lang.Validate;
import org.apache.poi.hssf.usermodel.HSSFFont;

public enum StyleUnderline {
    NONE("none", HSSFFont.U_NONE),
    SINGLE("single", HSSFFont.U_SINGLE),
    DOUBLE("double", HSSFFont.U_DOUBLE),
    SINGLE_ACCOUNTING("single accounting", HSSFFont.U_SINGLE_ACCOUNTING),
    DOUBLE_ACCOUNTING("double accounting", HSSFFont.U_DOUBLE_ACCOUNTING);

    private final String code;
    private final byte value;

    StyleUnderline(String code, byte value) {
        this.code = code;
        this.value = value;
    }

    public byte value() {
        return value;
    }

    public static StyleUnderline findByCode(String code) {
        Validate.notNull(code, "code");

        for (StyleUnderline alignment : values()) {
            if (code.equals(alignment.code)) {
                return alignment;
            }
        }

        return null;
    }
}
