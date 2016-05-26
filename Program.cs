// TTO Excel to CSV conversion
// Mike Tucker
// Github source: http://www.github.com/mtuckerinaz/ttoxls2csv
// mtucker6784@gmail.com
// http://www.tuckertechonline.com/

using System;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Text;

namespace XLSXtoCSV
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length >= 1)
            {
                if (args[0] == "-f")
                {
                    try
                    {
                        string[] lines;
                        var list = new List<string>();
                        var fileStream = new FileStream(args[1], FileMode.Open, FileAccess.Read);
                        using (var streamReader = new StreamReader(fileStream, Encoding.UTF8))
                        {
                            string line;
                            while ((line = streamReader.ReadLine()) != null)
                            {
                                list.Add(line);
                            }
                        }
                        lines = list.ToArray();
                        foreach(string s in lines)
                        {
                            string[] conv = s.Split(',');
                            if(File.Exists(conv[1]));
                            File.Delete(conv[1]);
                            conversion(conv[0], conv[1]);
                        }
                        Console.WriteLine("Conversion(s) was successful.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error: " + ex.Message.ToString() + Environment.NewLine);
                    }
                }
                else
                {
                    Microsoft.Office.Interop.Excel.Application conv2csv = new Microsoft.Office.Interop.Excel.Application();
                    try
                    {
                        conversion(args[0], args[1]);
                        Console.WriteLine("Conversion(s) was successful.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error: " + ex.Message.ToString() + Environment.NewLine);
                    }
                }

            }
            else
            {
                    Console.Write("No parameters given. Exiting." + Environment.NewLine);
            }

        }
        static void conversion(string input, string output)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application conv2csv = new Microsoft.Office.Interop.Excel.Application();
                Workbook wb = conv2csv.Workbooks.Open(input);
                wb.SaveAs(output, Microsoft.Office.Interop.Excel.XlFileFormat.xlCSVWindows, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlLocalSessionChanges, false, Type.Missing, Type.Missing, Type.Missing);
                wb.Close(false, Type.Missing, Type.Missing);
                conv2csv.Quit();

            }
            catch(Exception ex)
            {
                Console.Write(ex.Message.ToString());
            }
        }
    }
}
