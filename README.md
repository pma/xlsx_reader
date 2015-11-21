# XLSXReader

Reads XLSX using Apache POI and JInterface

## Installation

Download the following jars from http://mvnrepository.com/ and place them in the `priv` dir:

  * commons-codec-1.9.jar
  * commons-logging-1.1.3.jar
  * jinterface-1.5.6.jar
  * log4j-1.2.17.jar
  * poi-3.13-20150929.jar
  * poi-ooxml-3.13-20150929.jar
  * poi-ooxml-schemas-3.13-20150929.jar
  * xmlbeans-2.6.0.jar

Compile the java code with `javac -cp "priv/*" priv/XLSXReader.java`

Run with `iex --sname xlsx@localhost --cookie foo -S mix`

```iex
iex(xlsx@localhost)1> XLSXReader.parse "template.xlsx"
```
