package nl.gmt.excelBuilder;

import org.apache.commons.lang.Validate;
import org.apache.poi.hssf.util.HSSFColor;

public enum StyleColor {
    BLACK("black", HSSFColor.BLACK.index),
    BROWN("brown", HSSFColor.BROWN.index),
    OLIVE_GREEN("olive green", HSSFColor.OLIVE_GREEN.index),
    DARK_GREEN("dark green", HSSFColor.DARK_GREEN.index),
    DARK_TEAL("dark teal", HSSFColor.DARK_TEAL.index),
    DARK_BLUE("dark blue", HSSFColor.DARK_BLUE.index),
    INDIGO("indigo", HSSFColor.INDIGO.index),
    GREY_80_PERCENT("80%", HSSFColor.GREY_80_PERCENT.index),
    ORANGE("orange", HSSFColor.ORANGE.index),
    DARK_YELLOW("dark yellow", HSSFColor.DARK_YELLOW.index),
    GREEN("green", HSSFColor.GREEN.index),
    TEAL("teal", HSSFColor.TEAL.index),
    BLUE("blue", HSSFColor.BLUE.index),
    BLUE_GREY("blue grey", HSSFColor.BLUE_GREY.index),
    GREY_50_PERCENT("50%", HSSFColor.GREY_50_PERCENT.index),
    RED("red", HSSFColor.RED.index),
    LIGHT_ORANGE("light orange", HSSFColor.LIGHT_ORANGE.index),
    LIME("lime", HSSFColor.LIME.index),
    SEA_GREEN("sea green", HSSFColor.SEA_GREEN.index),
    AQUA("aqua", HSSFColor.AQUA.index),
    LIGHT_BLUE("light blue", HSSFColor.LIGHT_BLUE.index),
    VIOLET("violet", HSSFColor.VIOLET.index),
    GREY_40_PERCENT("40%", HSSFColor.GREY_40_PERCENT.index),
    PINK("pink", HSSFColor.PINK.index),
    GOLD("gold", HSSFColor.GOLD.index),
    YELLOW("yellow", HSSFColor.YELLOW.index),
    BRIGHT_GREEN("bright green", HSSFColor.BRIGHT_GREEN.index),
    TURQUOISE("turquoise", HSSFColor.TURQUOISE.index),
    DARK_RED("dark red", HSSFColor.DARK_RED.index),
    SKY_BLUE("sky blue", HSSFColor.SKY_BLUE.index),
    PLUM("plum", HSSFColor.PLUM.index),
    GREY_25_PERCENT("25%", HSSFColor.GREY_25_PERCENT.index),
    ROSE("rose", HSSFColor.ROSE.index),
    LIGHT_YELLOW("light yellow", HSSFColor.LIGHT_YELLOW.index),
    LIGHT_GREEN("light green", HSSFColor.LIGHT_GREEN.index),
    LIGHT_TURQUOISE("light turquoise", HSSFColor.LIGHT_TURQUOISE.index),
    PALE_BLUE("pale blue", HSSFColor.PALE_BLUE.index),
    LAVENDER("lavender", HSSFColor.LAVENDER.index),
    WHITE("white", HSSFColor.WHITE.index),
    CORNFLOWER_BLUE("cornflower blue", HSSFColor.CORNFLOWER_BLUE.index),
    LEMON_CHIFFON("lemon chiffon", HSSFColor.LEMON_CHIFFON.index),
    MAROON("maroon", HSSFColor.MAROON.index),
    ORCHID("orchid", HSSFColor.ORCHID.index),
    CORAL("coral", HSSFColor.CORAL.index),
    ROYAL_BLUE("royal blue", HSSFColor.ROYAL_BLUE.index),
    LIGHT_CORNFLOWER_BLUE("light cornflower blue", HSSFColor.LIGHT_CORNFLOWER_BLUE.index),
    TAN("tan", HSSFColor.TAN.index);

    private final String code;
    private final short value;

    StyleColor(String code, short value) {
        this.code = code;
        this.value = value;
    }

    public short value() {
        return value;
    }

    public static StyleColor findByCode(String code) {
        Validate.notNull(code, "code");

        for (StyleColor alignment : values()) {
            if (code.equals(alignment.code)) {
                return alignment;
            }
        }

        return null;
    }
}
