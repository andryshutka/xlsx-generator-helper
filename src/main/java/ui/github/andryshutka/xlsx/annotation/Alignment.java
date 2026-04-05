package ui.github.andryshutka.xlsx.annotation;

import org.apache.poi.ss.usermodel.HorizontalAlignment;

import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;

@Retention(RetentionPolicy.RUNTIME)
public @interface Alignment {
  HorizontalAlignment value() default HorizontalAlignment.GENERAL;
}
