import java.io.ObjectInputStream;
import java.net.URL;
import java.sql.ResultSetMetaData;
import java.util.zip.InflaterInputStream;
import javax.activation.DataHandler;
import javax.xml.namespace.QName;
import javax.sql.RowSet;
import com.epb.ap.EPBAP;
import com.epb.ap.EPBAPService;
import com.epb.ap.XProperties;
import com.epb.ap.XPropertiesEntry;

public class EPBReportQuery {
    public static void main(String[] args) throws Exception {
        if (args.length == 0) {
            throw new IllegalArgumentException("SQL argument is required");
        }
        String sql = args[0];
        int maxRows = args.length > 1 ? Integer.parseInt(args[1]) : 100000;

        EPBAPService service = new EPBAPService(
                new URL("http://192.168.1.177:8080/EPB_AP_EPB/EPB_AP?wsdl"),
                new QName("http://ap.epb.com/", "EPB_APService"));
        EPBAP port = service.getEPBAPPort();

        XProperties props = new XProperties();
        add(props, "dbId", "EPB");
        add(props, "timeZoneId", "Asia/Taipei");
        add(props, "preparedStatementSQL", sql);

        DataHandler handler = port.pullRowSetStream(props, null);
        int printed = 0;
        boolean headerPrinted = false;
        ObjectInputStream in = new ObjectInputStream(new InflaterInputStream(handler.getInputStream()));
        while (true) {
            Object object = in.readObject();
            if (object instanceof String) {
                String message = (String) object;
                if (message.length() > 0) {
                    System.err.println(message);
                }
                break;
            }
            RowSet rs = (RowSet) object;
            ResultSetMetaData md = rs.getMetaData();
            int cols = md.getColumnCount();
            if (!headerPrinted) {
                for (int i = 1; i <= cols; i++) {
                    if (i > 1) System.out.print("\t");
                    System.out.print(clean(md.getColumnLabel(i)));
                }
                System.out.println();
                headerPrinted = true;
            }
            while (rs.next()) {
                for (int i = 1; i <= cols; i++) {
                    if (i > 1) System.out.print("\t");
                    Object value = rs.getObject(i);
                    System.out.print(value == null ? "" : clean(value.toString()));
                }
                System.out.println();
                printed++;
                if (printed >= maxRows) {
                    System.err.println("MAX_ROWS_REACHED=" + maxRows);
                    return;
                }
            }
        }
    }

    private static String clean(String value) {
        return value.replace('\t', ' ').replace('\n', ' ').replace('\r', ' ');
    }

    private static void add(XProperties props, String key, Object value) {
        XPropertiesEntry entry = new XPropertiesEntry();
        entry.setKey(key);
        entry.setValue(value);
        props.getEntries().add(entry);
    }
}
