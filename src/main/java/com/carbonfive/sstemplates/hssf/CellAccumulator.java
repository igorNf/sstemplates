package com.carbonfive.sstemplates.hssf;

import java.util.ArrayList;
import java.util.List;
import java.util.Set;
import java.util.SortedMap;
import java.util.TreeMap;

//import org.apache.commons.lang.StringUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;

import org.apache.poi.ss.usermodel.Cell;

import com.carbonfive.sstemplates.Coordinates;


public class CellAccumulator {

    private SortedMap<Coordinates, Cell> cells;

    public CellAccumulator() {
        cells = new TreeMap<Coordinates, Cell>();
    }

    public void addCell (Cell cell, int row, int column) {
        cells.put(new Coordinates(row, (short) column), cell);
    }

    public String toString () {
        if (cells.size() == 1) {
            return cells.firstKey().toString();
        } else if (isContiguous(cells.keySet())) {
            return cells.firstKey() + ":" + cells.lastKey();
        } else {
            return StringUtils.join(cells.keySet().iterator(), ",");
        }
    }

    public String toString (String function, int max) {
        if (cells.size() == 1) {
            return function + "(" + cells.firstKey() + ")";
        } else if (isContiguous(cells.keySet())) {
            return function + "(" + cells.firstKey() + ":" + cells.lastKey() + ")";
        } else {
            return nestedCommaList(function, new ArrayList<Coordinates>(cells.keySet()), max);
        }
    }

    private static String nestedCommaList (String function, List<?> values, int max) {
        if (values.size() <= max)
            return function + "(" + StringUtils.join(values.iterator(), ",") + ")";

        List<String> list = new ArrayList<String>();
        for (int i = 0; i < values.size(); i += max) {
            list.add(nestedCommaList(function, values.subList(i, Math.min(i + max, values.size())),
                    max));
        }

        return nestedCommaList(function, list, max);
    }

    private boolean isContiguous (Set<Coordinates> coordinates) {
        int top = Integer.MAX_VALUE;
        int bottom = Integer.MIN_VALUE;
        short left = Short.MAX_VALUE;
        short right = Short.MIN_VALUE;

        for (Coordinates c : coordinates) {
            if (c.getRow() < top) top = c.getRow();
            if (c.getRow() > bottom) bottom = c.getRow();
            if (c.getColumn() < left) left = c.getColumn();
            if (c.getColumn() > right) right = c.getColumn();
        }
        return (((bottom - top + 1) * (right - left + 1)) == coordinates.size());
    }
}
