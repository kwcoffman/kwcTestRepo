/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package edu.umich.med.umms;

import com.sun.star.uno.XComponentContext;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;

/**
 *
 * @author kwc
 */
public class WatchDogAgain extends Thread {

    boolean opCompleted = false;
    int sleepTime = 30000;          // milliseconds, default is 30 seconds
    DecompUtil myutil = null;

    public WatchDogAgain() {
    }

    public void signalOpCompleted()
    {
        opCompleted = true;
    }

    public void setTimer(int timeInMilliseconds)
    {
        sleepTime = timeInMilliseconds;
    }

    @Override
    public void run()
    {
        System.err.println("WatchdogAgain " + this.getName() +
                ": starting up!");

        // Sleep for a while to let other operations proceed before checking again...
        try {
            System.err.println("WatchdogAgain " + this.getName() +
                    ": sleeping for " + (sleepTime/1000) + " seconds...");
            Thread.sleep(sleepTime);
        } catch (InterruptedException ex) {
            System.err.println("WatchdogAgain " + this.getName() +
                    ": interrupted exception!!");
            Logger.getLogger(WatchDogAgain.class.getName()).log(Level.SEVERE, null, ex);
        }

        if (opCompleted) {
            System.err.println("WatchdogAgain " + this.getName() +
                    ": the operation was completed, so my work is done!");
            return;
        }

        System.err.println("WatchdogAgain " + this.getName() +
                ": the operation did not complete in " + (sleepTime/1000) + " seconds!");

        KillOpenOffice killer = new KillOpenOffice();
        try {
            killer.killOpenOffice();
        } catch (IOException ex) {
            Logger.getLogger(WatchDogAgain.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
}
