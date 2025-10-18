# 
#  Purpose: Demonstrate usage of the ICVCommandEvents interface. This interface is used 
#           to send asynchronous notifications to the client application. 
# 
#           These notifications are:
#              OnLicenseError - Sent when a licensing error is detected in CODE V
#              OnCodeVError   - Sent when an error message is issued by CODE V
#              OnCodeVWarning - Sent when a warning message is issued by CODE V
#              OnPlotReady    - Sent when a plot file is ready to be displayed
# 
#           Note that all of these methods do not have to be implemented. Any events
#           that do not have an implemented sink method are discarded.
# Usage:
#           Run this file with the command:
#           python Example_CV_Events.py
#             

import sys
import shutil
from pythoncom import (CoInitializeEx, CoUninitialize, COINIT_MULTITHREADED, com_error )
from win32com.client import DispatchWithEvents
from win32api import FormatMessage

sys.coinit_flags = COINIT_MULTITHREADED
 
# ICVCommandEvents event sink class
class ICVCommandEvents:
    def OnLicenseError(self, error):
        # This event handler is called when a licensing error is 
        # detected in the CODE V application.
        print ("License error: %s " % error)

    def OnCodeVError(self, error):
        # This event handler is called when a CODE V error message is issued
        print ("CODE V error: %s " % error)

    def OnCodeVWarning(self, warning):
        # This event handler is called when a CODE V warning message is issued
        print ("CODE V warning: %s " % warning)

    def OnPlotReady(self, filename, plotwindow):
        # This event handler is called when a plot file, refered to by filename,
        # is ready to be displayed.
        # The event handler is responsible for saving/copying the
        # plot data out of the file specified by filename
        print ("CODE V Plot: %s in plot window %d" % (filename ,plotwindow) )
        shutil.copyfile(filename,"C:/CVUSER/myPlot.plt")
 
if __name__ == '__main__':

    try:
        #Create a CODE V Command object and connect to the ICVCommandEvents event sink
        cv = DispatchWithEvents("CodeV.Application", ICVCommandEvents)
        dir="c:\\cvuser"
        cv.StartingDirectory = dir
        cv.StartCodeV()
    
        # Generate some warning events
        result = cv.Command("res cv_lens:dbgauss")
        print(result)
    
        # Generate some plot events
        print("Each of these plots should go ")
        print("in the same plot window")
        print("---------------------------------------------")
        result = cv.Command("vie; go")
        print(result)
        print("---------------------------------------------")
        result = cv.Command("vie; go")
        print(result)
        print("---------------------------------------------")
        result = cv.Command("vie; go")
        print(result)
        print("---------------------------------------------")
        print("Each of these plots should go ")
        print("in the a different plot window")
        print("---------------------------------------------")
        result = cv.Command("wnd ope 3; wri 'Opening three plot windows'")
        print(result)
        print("---------------------------------------------")
        result = cv.Command("vie; go")
        print(result)
        print("---------------------------------------------")
        result = cv.Command("vie; go")
        print(result)
        print("---------------------------------------------")
        result = cv.Command("vie; go")
        print(result)
        print("---------------------------------------------")
    
        # Generate an error event
        result = cv.Command("in not_found 0")
        print(result)
        print("---------------------------------------------")
    
        cv.StopCodeV()
        # Delete the COM object
        del cv
    except com_error as error:
        print(error.strerror)
