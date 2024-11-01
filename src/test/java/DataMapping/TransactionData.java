package DataMapping;

import java.util.HashMap;
import java.util.Map;

public class TransactionData {
    private Map<String, String> rtData;
    private Map<String, String> dnData;

    public TransactionData() {
        this.rtData = new HashMap<>();
        this.dnData = new HashMap<>();
    }

    public void addRtData(String columnName, String value) {
        rtData.put(columnName, value);
    }

    public void addDnData(String columnName, String value) {
        dnData.put(columnName, value);
    }

    public String getRtData(String columnName) {
        return rtData.get(columnName);
    }

    public String getDnData(String columnName) {
        return dnData.get(columnName);
    }

    public Map<String, String> getRtDataMap() {
        return rtData;
    }

    public Map<String, String> getDnDataMap() {
        return dnData;
    }
}
