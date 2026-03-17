package rXlsx;

import java.util.Map;

public class series{
    int index;
    Map<String, Object> data;

    series(int index, Map<String, Object> data){
        this.index = index;
        this.data = data;
    }

    public void reindex(int newIndex){
        this.index = newIndex;
    }
}
