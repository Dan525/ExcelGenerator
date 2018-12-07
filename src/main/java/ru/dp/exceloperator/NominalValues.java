/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package ru.dp.exceloperator;

import java.util.HashMap;
import java.util.Map;

/**
 *
 * @author daniil_pozdeev
 */
public class NominalValues {
    
    private Map<Double, Integer> nominalMap;

    public NominalValues() {
        nominalMap = new HashMap<>();
    }

    public Map<Double, Integer> getNominalMap() {
        return nominalMap;
    }    
}
