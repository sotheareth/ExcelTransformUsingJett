# ExcelTransformUsingJett
Use JETT to transform excel

1. Go to https://svn.code.sf.net/p/jett/code-0/trunk/readme.txt
2. Implement small basic example
2. Understand tag foreach loop
3. That's all

//configuration add dependency maven or gradle below to pom.xml
//and all other dependency for jett
```sh
<dependency>
    <groupId>net.sf.jett</groupId>
    <artifactId>jett-core</artifactId>
    <version>0.10.0</version>
</dependency>
```

Apache POI 3.14 (or higher)
Apache Commons JEXL 2.1.1
Apache Commons Logging 1.2 (or higher)
SourceForge jAgg 0.9.0 (or higher)


// ----------------------create Excel template spreadsheet like below:

+----------------+----------------+
|${var}          |${var2}!        |
+----------------+----------------+

```sh
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
```

### EndUserPublicFlag class
```sh
public enum EndUserPublicFlag implements PersistableEnum<Integer> {
	
	PUBLIC_PRIVATE(0),  
	PRIVATE(1),
	PUBLIC(2);
	
	final int value;
	
	EndUserPublicFlag(int value) {
		this.value = value;
	}

	public Integer getValue() {
		return value;
	}
}
```

### convert class
```sh
import org.apache.commons.beanutils.ConversionException;

import com.bizmob.api.base.spring.jxls.JxlsConverter;
import com.bizmob.api.constant.enduser.EndUserPublicFlag;

public class EndUserPublicFlagConverter implements JxlsConverter {

	@SuppressWarnings("rawtypes")
	public Object convert( Class type, Object value) {
		if(value == null) {
			 throw new ConversionException("No value specified");
		}
		if(value instanceof String) {
			return EndUserPublicFlag.valueOf(value.toString());
		}
		return null;
	}

	public Class<?> getTargetClass() {
		return EndUserPublicFlag.class;
	}

}
```

### import data from excel
```sh
ConvertUtils.register(new EndUserPublicFlagConverter(), EndUserUseFlag.class);
		
		InputStream configXml = new FileInputStream(new File("D:/backup_old/Projects/portal-development/api-server/BIZMOB_HOME/admin/config/jxls/parsers/UserParser.xml"));
	    XLSReader mainReader = ReaderBuilder.buildFromXML(configXml);
	    
	    InputStream dataExcel = new FileInputStream(new File("D:/backup_old/Projects/portal-development/api-server/BIZMOB_HOME/admin/config/jxls/download/user_register_excel_template.xls"));
	    
	    Map<String, Object> beans = new HashMap<String, Object>();
	    List<InsertEndUserList> collection = new ArrayList<InsertEndUserList>();
	    
	    //array is same to config xml above
	    beans.put("array", collection);
	    mainReader.read(dataExcel, beans);
	    
	    System.out.println(beans);
```
