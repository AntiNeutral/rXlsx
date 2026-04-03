package rXlsx;

import java.util.LinkedHashMap;

public class series {
    public int index;
    public LinkedHashMap<String, Object> data;

    series(int index, LinkedHashMap<String, Object> data) {
        this.index = index;
        this.data = data;
    }

    public void reindex(int newIndex) {
        this.index = newIndex;
    }
}