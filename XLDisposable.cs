using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace financial_reporting
{
    class XLDisposable : IDisposable
    {
        object instance;
        bool   disposed = false;

        public XLDisposable(object instance)
        {
            this.instance = instance;
        }


        void IDisposable.Dispose()
        {
            if(disposed)
                return;

            try
            {
                     if(instance is Excel.Workbook)    ((Excel.Workbook   )instance).Close(0);
                else if(instance is Excel.Application) ((Excel.Application)instance).Quit();

                System.Runtime.InteropServices.Marshal.ReleaseComObject(instance);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                instance = null;
                disposed = true;
                GC.Collect();
                GC.Collect();
            }
        }
    }
}
