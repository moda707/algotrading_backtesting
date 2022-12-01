using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;

namespace BackTesting
{
    class ExcelReport
    {
        public void ExcelReportT_Overal(List<Chromosome> ChromList, int _StartRow, int _StartCol, ExcelPackage xlPackage, int Gener, string SymbolName)
        {

            ExcelWorksheet worksheetData;

            if (xlPackage.Workbook.Worksheets[SymbolName] == null)
            {
                xlPackage.Workbook.Worksheets.Add(SymbolName);
            }
            
            worksheetData = xlPackage.Workbook.Worksheets[SymbolName];
            

            #region Data
            int i = _StartRow;
            
            if(i>1 && !SymbolName.Contains("_")) worksheetData.SetValue(i - 1, 1, "Generation Number : " + Gener);

            foreach (var a in ChromList)
            {
                int j = 1;
                foreach (var w in ChromList[0].Genes)
                {
                    worksheetData.SetValue(i, j, a.Genes[j - 1]);
                    j++;
                }

                worksheetData.SetValue(i, j, a.J);//Profit Function
                j++;

                worksheetData.SetValue(i, j, a.SignalsNum);//Signals Number
                j++;

                i++;
            }

            //Summary
            //int j1 = 1;
            //var a1 = ChromList[0];
            //worksheetData1.SetValue(Gener + 3, j1, Gener + 1);
            //j1++;

            //foreach (var w in ChromList[0].Genes)
            //{
            //    worksheetData1.SetValue(Gener + 3, j1, a1.Genes[j1 - 2]);
            //    j1++;
            //}

            //worksheetData1.SetValue(Gener + 3, j1, a1.J);//Profit Function
            //j1++;

            //worksheetData1.SetValue(Gener + 3, j1, a1.SignalsNum);//Signals Number
            //j1++;

            #endregion
        }

        public void ExcelReport_Log(string strRep, ExcelPackage xlPackage, int Generation, string SymbolName)
        {
            ExcelWorksheet worksheetData;
            if (xlPackage.Workbook.Worksheets["Log"] == null)
            {
                xlPackage.Workbook.Worksheets.Add("Log");
            }

            worksheetData = xlPackage.Workbook.Worksheets["Log"];


            int i = 1;
            
            while (worksheetData.Cells[i, 1].Value != null)
            {
                i++;
            }

            worksheetData.SetValue(i, 1, SymbolName);
            worksheetData.SetValue(i, 2, Generation);
            worksheetData.SetValue(i, 3, strRep);

            
        }

        public void ExcelReport_General(string strRep, ExcelPackage xlPackage, int Generation, string SheetName)
        {
            ExcelWorksheet worksheetData;
            if (xlPackage.Workbook.Worksheets[SheetName] == null)
            {
                xlPackage.Workbook.Worksheets.Add(SheetName);
            }

            worksheetData = xlPackage.Workbook.Worksheets[SheetName];


            int i = 1;

            while (worksheetData.Cells[i, 1].Value != null)
            {
                i++;
            }

            
            worksheetData.SetValue(i, 1, strRep);
            worksheetData.SetValue(i, 2, Generation);
            worksheetData.SetValue(i, 3, strRep);
        }

        public void ExcelReport_ByChrom(Chromosome Chrome, Optimization OptFunc, ExcelPackage xlPackage, int Generation, string SheetName)
        {
            ExcelWorksheet worksheetData;
            if (xlPackage.Workbook.Worksheets[SheetName] == null)
            {
                xlPackage.Workbook.Worksheets.Add(SheetName);
            }

            worksheetData = xlPackage.Workbook.Worksheets[SheetName];


            int i = 1;

            while (worksheetData.Cells[i, 1].Value != null)
            {
                i++;
            }

            int k = 1;
            worksheetData.SetValue(i, k, OptFunc.Symb.Symbol);
            k++;

            worksheetData.SetValue(i, k, OptFunc.Indicator);
            k++;

            worksheetData.SetValue(i, k, OptFunc.Scope);
            k++;

            worksheetData.SetValue(i, k, OptFunc.Algorithm);
            k++;

            worksheetData.SetValue(i, k, OptFunc.SDEven);
            k++;
            worksheetData.SetValue(i, k, OptFunc.SHEven);
            k++;
            worksheetData.SetValue(i, k, OptFunc.EDEven);
            k++;
            worksheetData.SetValue(i, k, OptFunc.EHEven);
            k++;

            i++;

            k = 1;
            foreach (var a in Chrome.Genes)
            {
                worksheetData.SetValue(i, k, a);
                k++;
            }
            
            worksheetData.SetValue(i, k, Chrome.J);
            k++;

            worksheetData.SetValue(i, k, Chrome.SignalsNum);
        }
        
    }
}
