package glory.tools.excelutil.common;

import org.springframework.lang.NonNull;

public class LogStringUtil {

    public static final  String LOG_LINE      = "# =================================================================================================";
    private static final String SUB_LINE      = "# ==================== ";
    private static final String START         = "# ============================== ";
    private static final String END           = " ==============================";
    private static final int    TITLE_MAX_LEN = 35;

    private LogStringUtil() {
    }

    /**
     * Title
     *
     * @param value title
     */
    public static String logTitle(@NonNull String value) {
        final String text;
        int length = value.length();
        text = length <= TITLE_MAX_LEN
                ? value + " ".repeat(TITLE_MAX_LEN - length)
                : value.substring(0, TITLE_MAX_LEN);
        return START + text + END;
    }

    /**
     * Sub Title
     *
     * @param value sub title
     */
    public static String logSubTitle(@NonNull String value) {
        return SUB_LINE + value;
    }

}
