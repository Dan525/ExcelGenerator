/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package ru.dp.exceloperator;

import org.apache.poi.ss.util.CellAddress;

/**
 *
 * @author daniil_pozdeev
 */
public class TypeCursor {
    
    private CellAddress startAddress;
    private CellAddress endAddress;
    private int yCursor = 0;

    public CellAddress getStartAddress() {
        return startAddress;
    }

    public void setStartAddress(CellAddress startAddress) {
        this.startAddress = startAddress;
    }

    public CellAddress getEndAddress() {
        return endAddress;
    }

    public void setEndAddress(CellAddress endAddress) {
        this.endAddress = endAddress;
    }
    
    public void incrementCursor() {
        if (yCursor < endAddress.getRow()) {
            yCursor++;
        } else {
            throw new IndexOutOfBoundsException("Значение выходит за пределы таблицы");
        }
    }
}
