using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;

namespace VSTOAddInBradburyChart
{
    public partial class Ribbon
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
        }
        private void btnChartBradbury_Click(object sender, RibbonControlEventArgs e)
        {
            Worksheet currentSheet = Globals.ThisAddIn.GetActiveWorkSheet();
            var activeApp = Globals.ThisAddIn.GetActiveApp();

            int sRowCount = currentSheet.UsedRange.Rows.Count;
            Range serialNumbers = currentSheet.Range["A2:A" + sRowCount];
            Range dataCells = currentSheet.Range["B2:B" + sRowCount];

            // dataCells into array
            Object[,] originalData;
            originalData = dataCells.Value2;
            int elementCount = 0;
            double[] originalDataArr = new double[sRowCount - 1];
            string[] serialNumberArr = new string[sRowCount - 1];

            foreach (object element in originalData)
            {
                originalDataArr[elementCount] = Convert.ToDouble(element.ToString());
                elementCount++;
            }

            elementCount = 0;
            foreach (object element in serialNumberArr)
            {
                originalDataArr[elementCount] = Convert.ToDouble(element.ToString());
                elementCount++;
            }

            double standardDev = (double)activeApp.WorksheetFunction.StDevP(dataCells.Cells);
            double averageA = originalDataArr.Average();
            double threeStdDev = 3 * standardDev;

            //avg and std to worksheet
            string[] columnTitles = new[] { "Mean", "StdDev", "3StdDev" };
            string[] columnResults = new[] { averageA.ToString(), standardDev.ToString(), threeStdDev.ToString() };

            for (int n = 0; n<columnTitles.Length; n++)
            {
                Range titleRange = currentSheet.Range["D" + (n + 3)];
                titleRange.Value = columnTitles[n];
                Range valueRange = currentSheet.Range["E" + (n + 3)];
                valueRange.Value = columnResults[n];
            }
              }
            // make it look nicer !
            Range addBorder = currentSheet.Range["D3:E" + (columnTitles.Length + 2)];
            addBorder.BorderAround2(XlLineStyle.xlDouble, XlBorderWeight.xlMedium, XlColorIndex.xl
        }
    }
}
