package org.joget.marketplace;

import java.util.HashMap;

public class DuplicateAndSkip {
    private HashMap<String, Integer> map;
    private HashMap<String, Integer> skip;

    public DuplicateAndSkip() {
        map = new HashMap<>();
        skip = new HashMap<>();

    }

    public void setMap(HashMap<String, Integer> map) {
        this.map = map;
    }

    public Integer getSkipCount(String key) {
        if(skip.containsKey(key)) {
            return skip.get(key);
        }
        return 0;
    }

    public void addSkipCount(String key) {
        skip.put(key, skip.get(key) + 1);
    }

    public boolean checkKey(String key) {
        if (map.containsKey(key)) {
            if (!skip.containsKey(key)) {
                skip.put(key, 0);
            }
            return true;
        }
        return false;
    }

    public boolean skipCountLessThenDuplicate(String key) {
        if(map.get(key) > skip.get(key)) {
            return true;
        }
        return false;
    }
}
