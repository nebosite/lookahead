using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Reflection;

namespace LookAhead
{
    class OutlookHelper
    {
        public static Outlook.Application GetApplicationObject()
        {
            try
            {
                Outlook.Application application = null;

                // Check whether there is an Outlook process running.
                if (Process.GetProcessesByName("OUTLOOK").Count() > 0)
                {

                    // If so, use the GetActiveObject method to obtain the process and cast it to an Application object.
                    application = Marshal.GetActiveObject("Outlook.Application") as Outlook.Application;
                }
                else
                {

                    // If not, create a new instance of Outlook and log on to the default profile.
                    application = new Outlook.Application();
                    Outlook.NameSpace nameSpace = application.GetNamespace("MAPI");
                    nameSpace.Logon("", "", Missing.Value, Missing.Value);
                    nameSpace = null;
                }

                // Return the Outlook Application object.
                return application;
            }
            catch (COMException e)
            {
                if ((uint)e.ErrorCode == 0x800401E3)
                {
                    throw new ApplicationException("This user does not have rights to access the running outlook application.  \r\n" +
                        "Workarounds: Run under a different context (non-administrator)\r\n" +
                        "             Close Outlook (This will bring up a dialog box)\r\n" +
                        "             Turn off UAC.");
                }

                throw;
            }
        }

    }

}
