package nl.gmt.excelBuilder;

import org.apache.commons.lang.Validate;
import org.apache.poi.ss.usermodel.CellStyle;

public enum StyleFill {
    NONE("none", CellStyle.NO_FILL),
    SOLID("solid", CellStyle.SOLID_FOREGROUND),
    FINE_DOTS("fine dots", CellStyle.FINE_DOTS),
    ALT_BARS("alt bars", CellStyle.ALT_BARS),
    SPARSE_DOTS("sparse dots", CellStyle.SPARSE_DOTS),
    THICK_HORIZONTAL_BARS("thick horizontal bars", CellStyle.THICK_HORZ_BANDS),
    THICK_VERTICAL_BARS("thick vertical bars", CellStyle.THICK_VERT_BANDS),
    THICK_BACKWARD_DIAGONAL("thick backward diagonal", CellStyle.THICK_BACKWARD_DIAG),
    THICK_FORWARD_DIAGONAL("thick forward diagonal", CellStyle.THICK_FORWARD_DIAG),
    BIG_SPOTS("bit spots", CellStyle.BIG_SPOTS),
    BRICKS("bricks", CellStyle.BRICKS),
    THIN_HORIZONTAL_BANDS("thin horizontal bands", CellStyle.THIN_HORZ_BANDS),
    THIN_VERTICAL_BANDS("thin vertical bands", CellStyle.THIN_VERT_BANDS),
    THIN_BACKWARD_DIAGONAL("thin backward diagonal", CellStyle.THIN_BACKWARD_DIAG),
    THIN_FORWARD_DIAGONAL("thin forward diagonal", CellStyle.THIN_FORWARD_DIAG),
    SQUARES("squares", CellStyle.SQUARES),
    DIAMONDS("diamonds", CellStyle.DIAMONDS),
    LESS_DOTS("less dots", CellStyle.LESS_DOTS),
    LEAST_DOTS("least dots", CellStyle.LEAST_DOTS);

    private final String code;
    private final short value;

    StyleFill(String code, short value) {
        this.code = code;
        this.value = value;
    }

    public short value() {
        return value;
    }

    public static StyleFill findByCode(String code) {
        Validate.notNull(code, "code");

        for (StyleFill alignment : values()) {
            if (code.equals(alignment.code)) {
                return alignment;
            }
        }

        return null;
    }
}
