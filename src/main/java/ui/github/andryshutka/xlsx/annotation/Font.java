package ui.github.andryshutka.xlsx.annotation;

import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;

@Retention(RetentionPolicy.RUNTIME)
public @interface Font {
  String value() default "Arial";
  boolean bold() default false;
  boolean italic() default false;
  short size() default 8;
}
