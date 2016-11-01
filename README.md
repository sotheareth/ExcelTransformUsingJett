# ExcelTransformUsingJett
Use JETT to transform excel

1. Go to https://svn.code.sf.net/p/jett/code-0/trunk/readme.txt
2. Implement small basic example
2. Understand tag foreach loop
3. That's all

//configuration add dependency maven or gradle below to pom.xml
//and all other dependency for jett

<dependency>
    <groupId>net.sf.jett</groupId>
    <artifactId>jett-core</artifactId>
    <version>0.10.0</version>
</dependency>

Apache POI 3.14 (or higher)
Apache Commons JEXL 2.1.1
Apache Commons Logging 1.2 (or higher)
SourceForge jAgg 0.9.0 (or higher)


// ----------------------create Excel template spreadsheet like below:

+----------------+----------------+
|${var}          |${var2}!        |
+----------------+----------------+


//-------------------------use jett
Map<String, Object> beans = new HashMap<String, Object>();
beans.put("a", "Hello");
beans.put("b", "World");
ExcelTransformer transformer = new ExcelTransformer();
try
{
    transformer.transform("D:/backup_old/Projects/portal-development/api-server/BIZMOB_HOME/admin/config/jxls/template/template.xlsx", "D:/result.xlsx", beans);
}
catch (IOException e)
{
   System.err.println("I/O error occurred: " + e.getMessage());
}
catch (InvalidFormatException e)
{
   System.err.println("Spreadsheet was in invalid format: " + e.getMessage());
}
