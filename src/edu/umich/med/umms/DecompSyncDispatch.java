/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package edu.umich.med.umms;

import com.sun.star.beans.PropertyValue;
import com.sun.star.frame.DispatchResultEvent;
import com.sun.star.frame.DispatchResultState;
import com.sun.star.frame.XDesktop;
import com.sun.star.frame.XDispatch;
import com.sun.star.frame.XDispatchHelper;
import com.sun.star.frame.XDispatchProvider;
import com.sun.star.frame.XModel;
import com.sun.star.frame.XNotifyingDispatch;
import com.sun.star.frame.XSynchronousDispatch;
import com.sun.star.lang.EventObject;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.sun.star.util.URL;
import com.sun.star.util.XURLTransformer;
import java.util.concurrent.Semaphore;

/**
 *
 * @author kwc
 */
public class DecompSyncDispatch implements com.sun.star.frame.XDispatchResultListener {
    private Semaphore dispatchLock = new Semaphore(1, true);
    private com.spinn3r.log5j.Logger mylog = com.spinn3r.log5j.Logger.getLogger();
    private org.apache.log4j.Level myLogLevel = org.apache.log4j.Level.WARN;

    DecompSyncDispatch(org.apache.log4j.Level lvl)
    {
        mylog = com.spinn3r.log5j.Logger.getLogger();
        mylog.setLevel(lvl);
    }

    DecompSyncDispatch()
    {
        mylog = com.spinn3r.log5j.Logger.getLogger();
    }

    public void setLoggingLevel(org.apache.log4j.Level lvl)
    {
        myLogLevel = lvl;
        mylog.setLevel(myLogLevel);
    }

    /*
     * From http://www.oooforum.org/forum/viewtopic.phtml?t=71741
     */
    public void execSyncDispatch(XComponentContext xContext,
                                 XDesktop xDesktop,
                                 XMultiComponentFactory xMCF,
                                 Object pobjDoc,
                                 String cmd,
                                 PropertyValue[] props)
    {
        try {
            mylog.error("execSyncDispatch: ACQUIRING the lock (" + cmd + ")");
            dispatchLock.acquire();

            XMultiComponentFactory xFactory = (XMultiComponentFactory)
                    UnoRuntime.queryInterface(XMultiComponentFactory.class, xMCF);

            XModel xModel = (XModel) UnoRuntime.queryInterface(XModel.class, pobjDoc);
            XDispatchProvider xDispatchProvider = (XDispatchProvider)
                 UnoRuntime.queryInterface(XDispatchProvider.class, xModel.getCurrentController().getFrame());

            //Object object = xFactory.createInstance("com.sun.star.util.URLTransformer");
            Object object = xFactory.createInstanceWithContext("com.sun.star.util.URLTransformer", xContext);

            XURLTransformer xURLTransformer = (XURLTransformer) UnoRuntime.queryInterface(XURLTransformer.class, object);

            URL[] url = new URL[1];
            url[0] = new URL();
            url[0].Complete = cmd;
            xURLTransformer.parseStrict(url);

            // Request a dispatch object for the requested URL
            XDispatch xDispatcher = xDispatchProvider.queryDispatch(url[0], "", 0);
            if (xDispatcher == null) {
                mylog.error("execSyncDispatch: ***** Could not get XDispatch object! *****");
                // finally takes care of this?  dispatchLock.release();
                return;
            }

            XNotifyingDispatch xNotifyingDispatcher = (XNotifyingDispatch) UnoRuntime.queryInterface(XNotifyingDispatch.class, xDispatcher);
            if (xNotifyingDispatcher != null) {
                mylog.error("execSyncDispatch: (ASYNC?)  Dispatching command (" + cmd + ")");
                xNotifyingDispatcher.dispatchWithNotification(url[0], props, this);
                mylog.error("execSyncDispatch: RE-ACQUIRING the lock (" + cmd + ")");
                dispatchLock.acquire();
            } else {
                mylog.error("execSyncDispatch: NO XNotifyingDispatch for cmd (" + cmd + ")!!!");
                XDispatchProvider impressDispatchProvider = (XDispatchProvider)
                        UnoRuntime.queryInterface(XDispatchProvider.class, xModel.getCurrentController().getFrame());
                Object oDispatchHelper = xMCF.createInstanceWithContext("com.sun.star.frame.DispatchHelper", xContext);
                XDispatchHelper dispatchHelper = (XDispatchHelper) UnoRuntime.queryInterface(XDispatchHelper.class, oDispatchHelper);
                mylog.error("execSyncDispatch: (SYNC?) Dispatching command (" + cmd + ")");
                dispatchHelper.executeDispatch(impressDispatchProvider, cmd, "", 0, props);

            }
        } catch (java.lang.Exception e) {
            mylog.error("execSyncDispatch: Caught Exception: " + e.getMessage());
            e.printStackTrace();
        } finally {
            mylog.error("execSyncDispatch: RELEASING the lock (" + cmd + ")");
            dispatchLock.release();
            //DecompUtil.sleepFor(5);
        }
    }

    /*
     * From http://www.oooforum.org/forum/viewtopic.phtml?t=71741
     * XSynchronousDispatch is apparently not well-supported.
     */
    public void execSyncDispatch02(XComponentContext xContext,
                                 XDesktop xDesktop,
                                 XMultiComponentFactory xMCF,
                                 Object pobjDoc,
                                 String cmd,
                                 PropertyValue[] props)
    {
        try {
            dispatchLock.acquire();

            XMultiServiceFactory xFactory = (XMultiServiceFactory)
                    UnoRuntime.queryInterface(XMultiServiceFactory.class, xMCF);

            XModel xModel = (XModel) UnoRuntime.queryInterface(XModel.class, pobjDoc);
            XDispatchProvider xDispatchProvider = (XDispatchProvider)
                 UnoRuntime.queryInterface(XDispatchProvider.class, xModel.getCurrentController().getFrame());

            Object object = xFactory.createInstance("com.sun.star.util.URLTransformer");

            XURLTransformer xURLTransformer = (XURLTransformer) UnoRuntime.queryInterface(XURLTransformer.class, object);

            URL[] url = new URL[1];
            url[0] = new URL();
            url[0].Complete = cmd;
            xURLTransformer.parseStrict(url);

            // Request a dispatch object for the requested URL
            XDispatch xDispatcher = xDispatchProvider.queryDispatch(url[0], "", 0);

            // Alas, http://www.oooforum.org/forum/viewtopic.phtml?t=71617, says that most dispatch objects do not support this interface...
            XSynchronousDispatch xSynchronousDispatcher = (XSynchronousDispatch) UnoRuntime.queryInterface(XSynchronousDispatch.class, xDispatcher);
            if (xSynchronousDispatcher != null) {
                mylog.error("execSyncDispatch: Dispatching command: " + cmd);
                xSynchronousDispatcher.dispatchWithReturnValue(url[0], props);
            }
        } catch (java.lang.Exception e) {
            mylog.error("execSyncDispatch: Caught Exception: " + e.getMessage());
            e.printStackTrace();
        }
    }

    public void dispatchFinished(DispatchResultEvent e) {
        mylog.error("execSyncDispatch: The dispatch has \"finished\"!");
        switch (e.State) {
            case DispatchResultState.DONTKNOW:
            mylog.error("execSyncDispatch: The dispatch state is DONTKNOW");
            break;
            case DispatchResultState.FAILURE:
            mylog.error("execSyncDispatch: The dispatch state is FAILURE");
            break;
            case DispatchResultState.SUCCESS:
            mylog.error("execSyncDispatch: The dispatch state is SUCCESS");
            break;
        }
        dispatchLock.release();
    }

    public void disposing(EventObject e) {
        mylog.error("execSyncDispatch: ###### Huh? ###### The file is being disposed??? ###### : " + e.toString());
    }

}
