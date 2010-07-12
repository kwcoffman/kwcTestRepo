/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package edu.umich.med.umms;

/**
 *
 * @author kwc
 */
public class DecompException extends java.lang.Exception {

    private int exceptionError;

    DecompException(int errValue) {
        exceptionError = errValue;
    }

    DecompException(String errMsg) {
        super(errMsg);
    }

    @Override
    public String toString() {
        return "General DecompException[" + exceptionError + "]";
    }
}
