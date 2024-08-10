package model;

import java.util.Date;

public class Pair {
    String from;
    String to;

    public Pair(String from, String to) {
        this.from = from;
        this.to = to;
    }

    public String getFrom() {
        return from;
    }

    public String getTo() {
        return to;
    }

    @Override
    public String toString() {
        return from + " " + to + "\n";
    }
}
