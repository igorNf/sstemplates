package com.carbonfive.sstemplates;

@SuppressWarnings("unchecked")
public class Coordinates implements Comparable {

    private int   row;
    private short column;

    public int getRow () {
        return row;
    }

    public short getColumn () {
        return column;
    }

    public Coordinates(int row, short column) {
        this.row = row;
        this.column = column;
    }

    @Override
    public boolean equals (Object o) {
        if (o == null) return false;
        Coordinates p = (Coordinates) o;
        return ((row == p.getRow()) && (column == p.getColumn()));
    }

    @Override
    public int hashCode () {
        return row ^ column;
    }

    @Override
    public String toString () {
        return String.valueOf((char) (column + 65)) + (row + 1);
    }

    public int compareTo (Object o) {
        Coordinates c = (Coordinates) o;
        if (getRow() < c.getRow()) return -1;
        if (getRow() > c.getRow()) return 1;
        if (getColumn() < c.getColumn()) return -1;
        if (getColumn() > c.getColumn()) return 1;
        return 0;
    }
}
