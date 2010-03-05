/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package edu.umich.med.umms;

import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.frame.XDesktop;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import org.apache.commons.cli.Options;

/**
 *
 * @author kwc
 */
public class OpenOfficeUNODecompositionMain {
    
    private static final com.spinn3r.log5j.Logger mylog =
            com.spinn3r.log5j.Logger.getLogger(OpenOfficeUNODecomposition.class);

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args)
    {
        XComponentContext xContext = null;
        XMultiComponentFactory xMCF = null;
        XDesktop xDesktop = null;
        Options opt = new Options();
        DecompParameters dp = new DecompParameters();
        int err = 1;
        OpenOfficeUNODecomposition decomp = new OpenOfficeUNODecomposition();

        decomp.defineOptions(opt);
        err = decomp.processArguments(opt, args, dp);
        if (err != 0) {
            System.exit(1);
        }

        // get the remote office component context
        try {
            xContext = Bootstrap.bootstrap();
        } catch (com.sun.star.comp.helper.BootstrapException e) {
            mylog.fatal("Error connecting to OpenOffice process: " + e.getMessage());
            System.exit(3);
        }

        // get the remote office service manager
        xMCF = xContext.getServiceManager();


        try {
            /* A desktop environment contains tasks with one or more
            frames in which components can be loaded. Desktop is the
            environment for components which can instantiate within
            frames. */
            xDesktop = (XDesktop) UnoRuntime.queryInterface(XDesktop.class,
                    xMCF.createInstanceWithContext("com.sun.star.frame.Desktop",
                    xContext));
        } catch (com.sun.star.uno.Exception e) {
            mylog.fatal("Error getting OpenOffice desktop: " + e.getMessage());
            System.exit(4);
        }

        if (dp.getOperation() == DecompOperation.JSON) {
            err = decomp.processJsonFile(xContext, xDesktop, dp);
        } else {
            err = decomp.processSingleFile(xContext, xDesktop, dp);
        }

        System.exit(err);
    }
}
