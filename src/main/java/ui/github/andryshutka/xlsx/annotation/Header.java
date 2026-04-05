package ui.github.andryshutka.xlsx.annotation;

import org.apache.poi.ss.usermodel.IndexedColors;

import java.lang.annotation.Retention;
import java.lang.annotation.Target;

import static java.lang.annotation.ElementType.FIELD;
import static java.lang.annotation.RetentionPolicy.RUNTIME;
import static org.apache.poi.ss.usermodel.IndexedColors.GREY_25_PERCENT;

/**
 * Annotation for adjust header Excel cell <br/><br/>
 * label - header value or key from messages <br/>
 * widthSize - width of cell (ignored by default) <br/>
 * widthAsHeaderLength - default implementation of width <br/>
 * widthAsAverageColumnSize - width adjusted by average value length <br/>
 * color - colour of header <br/>
 * forCountry - set field only for adjusted country <br/>
 */
@Retention(RUNTIME)
@Target(FIELD)
public @interface Header {
    String label();
    boolean renderSummary() default false;
    int widthSize() default -1;
    boolean widthAsHeaderLength() default true;
    boolean widthAsAverageValueLength() default false;
    IndexedColors color() default GREY_25_PERCENT;
    String[] forCountry() default {};
}
