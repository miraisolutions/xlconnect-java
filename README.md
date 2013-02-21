xlconnect-java
==============

A simple Java wrapper around Apache POI, providing the backbone of R package [XLConnect](https://github.com/miraisolutions/xlconnect).

## Building

1. Build the Java package
```
    mvn package
```
2. Put the resulting JAR file into the `inst/java` directory of your `XLConnect` R package directory.
3. Install the R package (you'll likely want to restart your R session as well). 
