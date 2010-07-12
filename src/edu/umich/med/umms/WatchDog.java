/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package edu.umich.med.umms;

import com.sun.star.frame.XDesktop;
import com.sun.star.frame.XFrame;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.uno.XComponentContext;
import java.util.Timer;
import java.util.TimerTask;
import java.util.logging.Level;
import java.util.logging.Logger;

/**
 *
 * @author kwc
 */
public class WatchDog extends Thread {

    XComponentContext xContext;
    boolean timeToStop = false;
    DecompUtil myutil = null;

    public WatchDog(XComponentContext theContext) {
        xContext = theContext;
        myutil = new DecompUtil();
    }

    public void signalHalt()
    {
        timeToStop = true;
    }

    @Override
    public void run()
    {

        for (;;) {
            // Is it time to stop yet?
            if (timeToStop) {
                System.err.println("Watchdog " + this.getName() +
                        ": time to go byebye!");
                return;
            }

            System.err.println("Watchdog " + this.getName() +
                    ": nice to meet you!");
            // Must create a new Timer each time through!  If the timer gets
            // cancelled, it may not be re-used!
            Timer timer = new Timer(false);
            TimerTask killertask = new KillOpenOffice();


            // Schedule the timer
            System.err.println("Watchdog " + this.getName() +
                    ": scheduling a timer for 5 seconds from now...");
            timer.schedule(killertask, 5000);  // give it 5 seconds to complete

            // Attempt an operation
            try {
                System.err.println("Watchdog " + this.getName() +
                        ": trying to getServiceManager()");
                /*
                XMultiComponentFactory xMCF = xContext.getServiceManager();
                if (xMCF != null) {
                    System.err.println("Watchdog " + this.getName() + ": done with getServiceManager()");

                    String[] servicenames = xMCF.getAvailableServiceNames();
                    if (servicenames != null) {
                        System.err.println("Watchdog " + this.getName() + ": done with getAvailableServiceNames() found " + servicenames.length + " services!");
                        // If operation completes, OpenOffice is still alive, cancel timer
                        // (If timer goes off, that means we didn't get a response, so we
                        // need to kill Open Office and have it get restarted...)
                        System.err.println("Watchdog " + this.getName() + ": cancelling original timer task!!");
                        killertask.cancel();
                    }
                }*/
                XDesktop xDesktop = myutil.getDesktop(xContext);
                if (xDesktop != null) {
                    // If operation completes, OpenOffice is still alive, cancel timer
                    // (If timer goes off, that means we didn't get a response, so we
                    // need to kill Open Office and have it get restarted...)
                    System.err.println("Watchdog " + this.getName() +
                            ": got an XDesktop, cancelling original timer task!!");
                    killertask.cancel();
                }
            } catch (Exception ex) {
                System.err.println("Watchdog " + this.getName() +
                        ": caught exception: " + ex.getMessage());
                // Don't cancel the killertask in this case!
                timeToStop = true;
            }

            // Sleep for a while to let other operations proceed before checking again...
            try {
                System.err.println("Watchdog " + this.getName() +
                        ": sleeping for 10 seconds before starting again...");
                Thread.sleep(10000);
            } catch (InterruptedException ex) {
                System.err.println("Watchdog " + this.getName() +
                        ": interrupted exception!!");
                Logger.getLogger(WatchDog.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
    }
}
