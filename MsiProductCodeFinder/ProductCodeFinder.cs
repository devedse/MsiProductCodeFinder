using System;
using System.Text;
using System.Runtime.InteropServices;

namespace MsiProductCodeFinder
{
    class ProductCodeFinder
    {
        [DllImport("msi.dll", SetLastError = true)]
        static extern uint MsiOpenDatabase(string szDatabasePath, Int32 phPersist, out Int32 phDatabase);

        [DllImport("msi.dll", CharSet = CharSet.Unicode)]
        static extern int MsiDatabaseOpenViewW(Int32 hDatabase, [MarshalAs(UnmanagedType.LPWStr)] string szQuery, out Int32 phView);

        [DllImport("msi.dll", CharSet = CharSet.Unicode)]
        static extern int MsiViewExecute(Int32 hView, Int32 hRecord);

        [DllImport("msi.dll", CharSet = CharSet.Unicode)]
        static extern uint MsiViewFetch(Int32 hView, out Int32 hRecord);

        [DllImport("msi.dll", CharSet = CharSet.Unicode)]
        static extern int MsiRecordGetString(Int32 hRecord, int iField,
           [Out] StringBuilder szValueBuf, ref int pcchValueBuf);

        [DllImport("msi.dll", ExactSpelling = true)]
        static extern Int32 MsiCreateRecord(uint cParams);

        [DllImport("msi.dll", ExactSpelling = true)]
        static extern uint MsiCloseHandle(Int32 hAny);

        public static void WriteErrorIfNotEqual(string methodName, uint expected, uint actual, StringBuilder errorBuilder)
        {
            if (actual != expected)
            {
                errorBuilder.AppendLine($"{methodName} returned exit code: {actual}, Expected: {expected}.");
            }
        }

        public static void WriteErrorIfNotEqual(string methodName, int expected, int actual, StringBuilder errorBuilder)
        {
            if (actual != expected)
            {
                errorBuilder.AppendLine($"{methodName} returned exit code: {actual}, Expected: {expected}.");
            }
        }

        public static string GetVersionInfo(string fileName, out StringBuilder errorBuilder)
        {
            errorBuilder = new StringBuilder();
            errorBuilder.AppendLine();

            string sqlStatement = "SELECT * FROM Property WHERE Property = 'ProductCode'";

            int phDatabase = 0;
            int phView = 0;
            int hRecord = 0;

            StringBuilder szValueBuf = new StringBuilder();
            //Lenght of variable 
            szValueBuf.Capacity = 33;
            int pcchValueBuf = 255;

            try
            {
                // Open the MSI database in the input file 
                try
                {
                    uint val = MsiOpenDatabase(fileName, 0, out phDatabase);
                    WriteErrorIfNotEqual("MsiOpenDatabase", 0, val, errorBuilder);
                }
                catch (Exception ex)
                {
                    // Add useful information to the exception
                    throw new ApplicationException("Something wrong happened during opening MSI :", ex);
                }

                hRecord = MsiCreateRecord(3);

                // Open a view on the Property table for the version property 
                try
                {
                    int viewVal = MsiDatabaseOpenViewW(phDatabase, sqlStatement, out phView);
                    WriteErrorIfNotEqual("MsiDatabaseOpenViewW", 0, viewVal, errorBuilder);
                }
                catch (Exception ex)
                {
                    // Add useful information to the exception
                    throw new ApplicationException("Something wrong happened during opening Property table:", ex);
                }

                // Execute the view query 
                int exeVal = MsiViewExecute(phView, hRecord);
                WriteErrorIfNotEqual("MsiViewExecute", 0, exeVal, errorBuilder);

                // Get the record from the view 
                uint fetchVal = MsiViewFetch(phView, out hRecord);
                WriteErrorIfNotEqual("MsiViewFetch", 0, fetchVal, errorBuilder);

                // Get the version from the data 
                int retVal = MsiRecordGetString(hRecord, 2, szValueBuf, ref pcchValueBuf);
                WriteErrorIfNotEqual("MsiRecordGetString", 0, retVal, errorBuilder);
            }
            finally
            {
                uint uRetCode;
                uRetCode = MsiCloseHandle(phDatabase);
                uRetCode = MsiCloseHandle(phView);
                uRetCode = MsiCloseHandle(hRecord);
            }
            return szValueBuf.ToString();
        }

        //public static string GetVersionInfoWIX(string fileName)
        //{
        //    string szProductCode;

        //    using (var database = new QDatabase(fileName, DatabaseOpenMode.ReadOnly))
        //    {
        //        szProductCode = database.ExecutePropertyQuery("ProductCode");
        //    }
        //    return szProductCode;
        //}

        public static ProductCodeResult ObtainProductCode(string strFileName)
        {
            try
            {
                //Check if MSI file exists
                if (System.IO.File.Exists(strFileName) == true)
                {
                    StringBuilder errorBuilder;
                    var strGUID = GetVersionInfo(strFileName, out errorBuilder);

                    if (string.IsNullOrEmpty(strGUID))
                    {
                        return new ProductCodeResult(false, $"Something went wrong. ProductCode wasn't obtained. Errors: {errorBuilder.ToString()}");
                    }
                    else
                    {
                        return new ProductCodeResult(true, strGUID.Trim());
                    }
                }
                else
                {
                    return new ProductCodeResult(false, "The specified path doesn't exist.");
                }
            }
            catch (Exception ex)
            {
                return new ProductCodeResult(false, ex.ToString());
            }
        }
    }
}