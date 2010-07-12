/*
 * From the OpenOffice Forum (www.oooforum.org)
 * This class allows you to kill instances of OpenOffice.
 */
package edu.umich.med.umms;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.TimerTask;

/**
 *
 * @author hol.sten (http://www.oooforum.org/forum/viewtopic.phtml?t=5019&start=16)
 */
public class KillOpenOffice extends TimerTask {

    /**
     * Kill OpenOffice.
     *
     * Supports Windows XP, Solaris (SunOS) and Linux.
     * Perhaps it supports more OS, but it has been tested
     * only with this three.
     *
     * @throws IOException  Killing OpenOffice didn't work
     */
    public static void killOpenOffice() throws IOException {
        String osName = System.getProperty("os.name");
        if (osName.startsWith("Windows")) {
            windowsKillOpenOffice();
        } else if (osName.startsWith("SunOS") || osName.startsWith("Linux")) {
            System.err.println("KillOpenOffice: doing unix kill");
            unixKillOpenOffice();
        } else {
            throw new IOException("Unknown OS, killing impossible");
        }
    }

    /**
     * Kill OpenOffice on Windows XP.
     */
    private static void windowsKillOpenOffice() throws IOException {
        Runtime.getRuntime().exec("tskill soffice");
    }

    /**
     * Kill OpenOffice on Unix.
     */
    private static void unixKillOpenOffice() throws IOException {
        Runtime runtime = Runtime.getRuntime();

        String pid = getOpenOfficeProcessID();
        while (pid != null) {
            System.err.println("KillOpenOffice: killing pid: " + pid);
            String[] killCmd = {"/bin/bash", "-c", "kill -9 " + pid};
            runtime.exec(killCmd);

            // Is another OpenOffice prozess running?
            pid = getOpenOfficeProcessID();
        }
    }

    /**
     * Get OpenOffice prozess id.
     */
    private static String getOpenOfficeProcessID() throws IOException {
        Runtime runtime = Runtime.getRuntime();

        // Get prozess id
        String[] getPidCmd = {"/bin/bash", "-c", "ps -e|grep soffice|awk '{print $1}'"};
        Process getPidProcess = runtime.exec(getPidCmd);

        // Read prozess id
        InputStreamReader isr = new InputStreamReader(getPidProcess.getInputStream());
        BufferedReader br = new BufferedReader(isr);

        return br.readLine();
    }

    @Override
    public void run() {
        try {
            System.err.println("KillOpenOffice: about to actually kill openoffice!");
            KillOpenOffice.killOpenOffice();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
