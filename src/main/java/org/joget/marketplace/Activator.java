package org.joget.marketplace;

import java.util.ArrayList;
import java.util.Collection;
import org.osgi.framework.BundleActivator;
import org.osgi.framework.BundleContext;
import org.osgi.framework.ServiceRegistration;

public class Activator implements BundleActivator {

    public static final String VERSION = "8.0.15";
    
    protected Collection<ServiceRegistration> registrationList;

    public void start(BundleContext context) {
        registrationList = new ArrayList<ServiceRegistration>();

        //Register plugin here
        registrationList.add(context.registerService(DownloadCsvOrExcelDatalistAction.class.getName(), new DownloadCsvOrExcelDatalistAction(), null));
        registrationList.add(context.registerService(DownloadCsvOrExcelTool.class.getName(), new DownloadCsvOrExcelTool(), null));
    }

    public void stop(BundleContext context) {
        for (ServiceRegistration registration : registrationList) {
            registration.unregister();
        }
    }
}