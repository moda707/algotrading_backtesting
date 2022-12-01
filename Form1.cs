using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using vtocSqlInterface;
using Accord;
using OfficeOpenXml;
using System.IO;

namespace BackTesting
{
    public partial class Form1 : Form
    {
        private List<Symbols> Symbol;
        private List<Symbols> FinalSymbol;
        
        private vtocSqlInterface.sqlInterface mySql;
        private string sqlCmd;
        private int Scop;
        public DataTable prow = new DataTable();
        public double[] Totalcapital;
        public Int32 InitialCash = 0;
        public Int32 InitialShare = 0;
        public Int32 DefaultVol = 100;
        public string [,] ScopList;
        public List<double> range;
        public List<List<double>> ChromRange;
        public Random RVar = new Random();
        
        private string OutputFileName;
        FileInfo newFile;
        FileInfo newFileT;

        ExcelPackage xlPackage;
        ExcelPackage xlPackageT;

        private delegate void SetTextCallback(System.Windows.Forms.Control control, string text);

        public Form1()
        {
            InitializeComponent();
            

            Symbol = new List<Symbols>();
            FinalSymbol = new List<Symbols>();

            ReadSymbols();
            Symbol = Symbol.OrderBy(t => t.Symbol).ToList();
            lstFirstSymbol.DataSource = Symbol;
            lstFirstSymbol.ValueMember = "InsCode";
            lstFirstSymbol.DisplayMember = "Symbol";  
        }

        private void ReadSymbols()
        {
            mySql = new sqlInterface(Properties.Settings.Default.sqlserver, "AdjPrice",
                                     Properties.Settings.Default.username, Properties.Settings.Default.pass);
            sqlCmd = @"  SELECT DISTINCT S.LVal18AFC as SymbolName, S.InsCode as Symbol,S.LVal30
                          FROM [TseTrade].[dbo].[vwTseInstrument] S
                          JOIN TseTrade.dbo.vwTsePrice T ON T.InsCode = S.InsCode
                          WHERE S.Flow in (1,2) and YMarNSC='No' and YVal in (300 ,303)
                          ORDER BY LVal18AFC";
            DataTable dtSymbols = mySql.SqlExecuteReader(sqlCmd);
            foreach (DataRow row in dtSymbols.Rows)
            {
                if (dtSymbols.Columns.Contains("Symbol") && dtSymbols.Columns.Contains("SymbolName"))
                {
                    var sCode = row["Symbol"].ToString();
                    var sName = row["SymbolName"].ToString();

                    Symbol.Add(new Symbols(sName, sCode));
                }
            }
        }
        
        private void btnoneright_Click(object sender, EventArgs e)
        {
            lstFinalSymbol.Items.Add((Symbols)lstFirstSymbol.SelectedItem);
            lstFinalSymbol.ValueMember = "InsCode";
            lstFinalSymbol.DisplayMember = "Symbol";
            lstFinalSymbol.SelectedIndex = lstFinalSymbol.Items.Count-1;

            FinalSymbol.Add((Symbols)lstFirstSymbol.SelectedItem);
        }

        private void btnoneleft_Click(object sender, EventArgs e)
        {
            FinalSymbol.Remove((Symbols)lstFinalSymbol.SelectedItem);
            lstFinalSymbol.Items.Remove((Symbols)lstFinalSymbol.SelectedItem);            
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            cmbalg.SelectedIndex = 0;
            cmbScop.SelectedIndex = 0;
            cmbOptAlg.SelectedIndex = 0;
            DateTime NTime = DateTime.Now;

            txtDEven.Text = NTime.Year.ToString();

            if (NTime.Month.ToString().Length == 1)
            {
                txtDEven.Text += "0" + NTime.Month.ToString();
            }
            else
            {
                txtDEven.Text += NTime.Month.ToString();
            }
            if (NTime.Day.ToString().Length == 1)
            {
                txtDEven.Text += "0" + NTime.Day.ToString();
            }
            else
            {
                txtDEven.Text += NTime.Day.ToString();
            }

            txtHEven.Text = NTime.Hour.ToString();
            if (NTime.Minute.ToString().Length == 1)
            {
                txtHEven.Text += "0" + NTime.Minute.ToString();
            }
            else
            {
                txtHEven.Text += NTime.Minute.ToString();
            }

            if (NTime.Second.ToString().Length == 1)
            {
                txtHEven.Text += "0" + NTime.Second.ToString();
            }
            else
            {
                txtHEven.Text += NTime.Second.ToString();
            }
        }

        private void txtsymbolsearch_TextChanged(object sender, EventArgs e)
        {
            List<Symbols> tmpSymbol = new List<Symbols>();

            tmpSymbol = Symbol.FindAll(t => t.Symbol.Contains(txtsymbolsearch.Text));
            lstFirstSymbol.DataSource = tmpSymbol;
            lstFirstSymbol.ValueMember = "InsCode";
            lstFirstSymbol.DisplayMember = "Symbol";
            lstFirstSymbol.SelectedValue = "Selected";
        }

        private void btnTrade_Click(object sender, EventArgs e)
        {

            //var smb = (Symbols)lstFinalSymbol.SelectedItem;
            //Totalcapital = new double[Convert.ToInt16(txttestperiod.Text)];
            //backgroundWorker1.RunWorkerAsync();
            //int TestPeriod = Convert.ToInt16(txttestperiod.Text);
            //int MPeriod = Convert.ToInt16(lstperiod.Text);
            //double Cash = Convert.ToDouble(txtinitcash.Text);
            //double Share = Convert.ToDouble(txtinitshare.Text);
            //Int32 DEven = Convert.ToInt32(txtDEven.Text);
            //Int32 HEven = Convert.ToInt32(txtHEven.Text);

            //switch (cmbScop.Text)
            //{
            //    case "Daily":
            //        Scop = 3600;
            //        break;
            //    case "1 Hour":
            //        Scop = 60;
            //        break;
            //    case "10 Minute":
            //        Scop = 10;
            //        break;
            //    case "1 Minute":
            //        Scop = 1;
            //        break;
            //}

            ////DataTable BestLimits = mySql.SqlExecuteReader(sqlCmd);

            //progressBar1.Value = 5;
            //sqlCmd = @"SELECT * FROM (SELECT TOP (" + TestPeriod + ") T.DEven, T.HEven, T.price, ((CASE WHEN T.Buy>0 THEN 1 ELSE 0 END) + (CASE WHEN T.Sell>0 THEN -1 ELSE 0 END))AlligatorFS,((CASE WHEN RSI.Buy>0 THEN 1 ELSE 0 END) + (CASE WHEN RSI.Sell>0 THEN -1 ELSE 0 END))RSIFS, ((CASE WHEN CCI.Buy>0 THEN 1 ELSE 0 END) + (CASE WHEN CCI.Sell>0 THEN -1 ELSE 0 END))CCIFS, ((CASE WHEN MACD.Buy>0 THEN 1 ELSE 0 END) + (CASE WHEN MACD.Sell>0 THEN -1 ELSE 0 END))MACDFS, ISNULL((SELECT TOP 1 T2.PMeDem FROM [TseTrade].[dbo].[TseBestLimits] T2 WHERE T2.InsCode = " + smb.InsCode + " AND T2.number=1 AND ((" + Scop + "=1 AND T2.DEven > T.DEven) or (" + Scop + " in (2,3,4) AND T2.DEven=T.DEVEN AND T2.HEven>T.HEVEN))),0) PMeDem, ISNULL((SELECT TOP 1 T2.PMeOf FROM [TseTrade].[dbo].[TseBestLimits] T2 WHERE T2.InsCode = " + smb.InsCode + " AND T2.number=1 AND ((" + Scop + "=1 AND T2.DEven > T.DEven) or (" + Scop + " in (2,3,4) AND T2.DEven=T.DEVEN AND T2.HEven>T.HEVEN))),0) PMeOf, ISNULL((SELECT TOP 1 T2.QTitMeDem FROM [TseTrade].[dbo].[TseBestLimits] T2 WHERE T2.InsCode = " + smb.InsCode + " AND T2.number=1 AND ((" + Scop + "=1 AND T2.DEven > T.DEven) or (" + Scop + " in (2,3,4) AND T2.DEven=T.DEVEN AND T2.HEven>T.HEVEN))),0) QTitMeDem, ISNULL((SELECT TOP 1 T2.QTitMeOf FROM [TseTrade].[dbo].[TseBestLimits] T2 WHERE T2.InsCode = " + smb.InsCode + " AND T2.number=1 AND ((" + Scop + "=1 AND T2.DEven > T.DEven) or (" + Scop + " in (2,3,4) AND T2.DEven=T.DEVEN AND T2.HEven>T.HEVEN))),0) QTitMeOf FROM TseTrade.dbo.fn_AT_IND_Alligator_SIGNALS(" + smb.InsCode + ",13,8,5,8,5,3,0.1," + MPeriod + "," + Scop + ",1000," + DEven + "," + HEven + ") T JOIN TseTrade.dbo.fn_AT_IND_RSI_SIGNALS(" + smb.InsCode + ",14," + MPeriod + ",70,30," + Scop + ",1000," + DEven + "," + HEven + ") RSI ON RSI.DEven = T.DEven AND RSI.HEven = T.HEven JOIN TseTrade.dbo.fn_AT_IND_CCI_SIGNALS(" + smb.InsCode + ",14,100,-100,0.015," + Scop + ",1000," + DEven + "," + HEven + ") CCI ON CCI.DEven = T.DEven AND CCI.HEven = T.HEven JOIN TseTrade.dbo.fn_AT_IND_MACD_SIGNALS(" + smb.InsCode + ",26,12,9," + MPeriod + "," + Scop + ",1000," + DEven + "," + HEven + ") MACD ON MACD.DEven = T.DEven AND MACD.HEven = T.HEven ORDER BY T.DEven DESC,T.HEven DESC) J ORDER BY J.DEven, J.HEven";
                        
            //DataTable ALLSignals = mySql.SqlExecuteReader(sqlCmd);
            
            //progressBar1.Value = 60;


            //DataTable AlgFlow = new DataTable();
            //AlgFlow.Columns.Add("DEven");
            //AlgFlow.Columns.Add("HEven");
            //AlgFlow.Columns.Add("price");
            //AlgFlow.Columns.Add("AlligatorFS");
            //AlgFlow.Columns.Add("ACash");
            //AlgFlow.Columns.Add("AShare");
            //AlgFlow.Columns.Add("ARate");
            //AlgFlow.Columns.Add("RSIFS");
            //AlgFlow.Columns.Add("RCash");
            //AlgFlow.Columns.Add("RShare");
            //AlgFlow.Columns.Add("RRate");
            //AlgFlow.Columns.Add("CCIFS");
            //AlgFlow.Columns.Add("CCash");
            //AlgFlow.Columns.Add("CShare");
            //AlgFlow.Columns.Add("CRate");
            //AlgFlow.Columns.Add("MACDFS");            
            //AlgFlow.Columns.Add("MCash");
            //AlgFlow.Columns.Add("MShare");
            //AlgFlow.Columns.Add("MRate");
            //AlgFlow.Columns.Add("Combine");
            //AlgFlow.Columns.Add("ComCash");
            //AlgFlow.Columns.Add("ComShare");
            //AlgFlow.Columns.Add("ComRate");
            //AlgFlow.Columns.Add("B-H Rate");
            //AlgFlow.Rows.Add(0,0,0,0,Cash, Share, 0,0,Cash, Share, 0,0,Cash, Share, 0,0,Cash, Share, 0,0,Cash, Share, 0,0);

            //progressBar1.Value = 65;

            //int i =1;
            //foreach (DataRow a in ALLSignals.Rows)
            //{
            //    double[] tmpcash = { 0, 0, 0, 0, 0 };//0:Alligator,1:RSI,2:CCI,3:MACD,4:Combine
            //    double[] tmpshare = { 0, 0, 0, 0, 0 };//0:Alligator,1:RSI,2:CCI,3:MACD,4:Combine
            //    double[] tmprate = { 0, 0, 0, 0, 0 , 0 };//0:Alligator,1:RSI,2:CCI,3:MACD,4:Combine,5:Buy and Hold
            //    //Alligator
            //    if (Convert.ToInt32(a["AlligatorFS"]) > 0)
            //    {
            //        tmpshare[0] = Math.Min((Int32)(0.1 * Convert.ToDouble(AlgFlow.Rows[i - 1]["ACash"]) / Convert.ToDouble(a["PMeOf"])), Convert.ToInt64(a["QTitMeOf"]));
            //        tmpcash[0] = -1.005 * Convert.ToDouble(a["PMeOf"]) * tmpshare[0];
            //    }
            //    else if (Convert.ToInt32(a["AlligatorFS"]) < 0)
            //    {
            //        tmpshare[0] = -Math.Min(Convert.ToInt32(AlgFlow.Rows[i - 1]["AShare"]), Convert.ToInt64(a["QTitMeDem"]));
            //        tmpcash[0] = 0.99 * (-tmpshare[0]) * Convert.ToDouble(a["PMeDem"]);
            //    }

            //    //RSI
            //    if (Convert.ToInt32(a["RSIFS"]) > 0)
            //    {
            //        tmpshare[1] = Math.Min((Int32)(0.1 * Convert.ToDouble(AlgFlow.Rows[i - 1]["RCash"]) / Convert.ToDouble(a["PMeOf"])), Convert.ToInt64(a["QTitMeOf"]));
            //        tmpcash[1] = -1.005 * Convert.ToDouble(a["PMeOf"]) * tmpshare[1];
            //    }
            //    else if (Convert.ToInt32(a["RSIFS"]) < 0)
            //    {
            //        tmpshare[1] = -Math.Min(Convert.ToInt32(AlgFlow.Rows[i - 1]["RShare"]), Convert.ToInt64(a["QTitMeDem"]));
            //        tmpcash[1] = 0.99 * (-tmpshare[1]) * Convert.ToDouble(a["PMeDem"]);
            //    }

            //    //CCI
            //    if (Convert.ToInt32(a["CCIFS"]) > 0)
            //    {
            //        tmpshare[2] = Math.Min((Int32)(0.1 * Convert.ToDouble(AlgFlow.Rows[i - 1]["CCash"]) / Convert.ToDouble(a["PMeOf"])), Convert.ToInt64(a["QTitMeOf"]));
            //        tmpcash[2] = -1.005 * Convert.ToDouble(a["PMeOf"]) * tmpshare[2];
            //    }
            //    else if (Convert.ToInt32(a["CCIFS"]) < 0)
            //    {
            //        tmpshare[2] = -Math.Min(Convert.ToInt32(AlgFlow.Rows[i - 1]["CShare"]), Convert.ToInt64(a["QTitMeDem"]));
            //        tmpcash[2] = 0.99 * (-tmpshare[2]) * Convert.ToDouble(a["PMeDem"]);
            //    }

            //    //MACD
            //    if (Convert.ToInt32(a["MACDFS"]) > 0)
            //    {
            //        tmpshare[3] = Math.Min((Int32)(0.1 * Convert.ToDouble(AlgFlow.Rows[i - 1]["MCash"]) / Convert.ToDouble(a["PMeOf"])), Convert.ToInt64(a["QTitMeOf"]));
            //        tmpcash[3] = -1.005 * Convert.ToDouble(a["PMeOf"]) * tmpshare[3];
            //    }
            //    else if (Convert.ToInt32(a["MACDFS"]) < 0)
            //    {
            //        tmpshare[3] = -Math.Min(Convert.ToInt32(AlgFlow.Rows[i - 1]["MShare"]), Convert.ToInt64(a["QTitMeDem"]));
            //        tmpcash[3] = 0.99 * (-tmpshare[3]) * Convert.ToDouble(a["PMeDem"]);
            //    }

            //    int TotalSignals = Convert.ToInt32(a["AlligatorFS"]) + Convert.ToInt32(a["RSIFS"]) + Convert.ToInt32(a["CCIFS"]) + Convert.ToInt32(a["MACDFS"]);

            //    //Combine
            //    if (TotalSignals >= 2)
            //    {
            //        tmpshare[4] = Math.Min((Int32)(0.1 * Convert.ToDouble(AlgFlow.Rows[i - 1]["ComCash"]) / Convert.ToDouble(a["PMeOf"])), Convert.ToInt64(a["QTitMeOf"]));
            //        tmpcash[4] = -1.005 * Convert.ToDouble(a["PMeOf"]) * tmpshare[4];
            //    }
            //    else if (TotalSignals <= -2)
            //    {
            //        tmpshare[4] = -Math.Min(Convert.ToInt32(AlgFlow.Rows[i - 1]["ComShare"]), Convert.ToInt64(a["QTitMeDem"]));
            //        tmpcash[4] = 0.99 * (-tmpshare[4]) * Convert.ToDouble(a["PMeDem"]);
            //    }

            //    tmpcash[0] += Convert.ToDouble(AlgFlow.Rows[i - 1]["ACash"]);
            //    tmpshare[0] += Convert.ToInt32(AlgFlow.Rows[i - 1]["AShare"]);

            //    tmpcash[1] += Convert.ToDouble(AlgFlow.Rows[i - 1]["RCash"]);
            //    tmpshare[1] += Convert.ToInt32(AlgFlow.Rows[i - 1]["RShare"]);

            //    tmpcash[2] += Convert.ToDouble(AlgFlow.Rows[i - 1]["CCash"]);
            //    tmpshare[2] += Convert.ToInt32(AlgFlow.Rows[i - 1]["CShare"]);

            //    tmpcash[3] += Convert.ToDouble(AlgFlow.Rows[i - 1]["MCash"]);
            //    tmpshare[3] += Convert.ToInt32(AlgFlow.Rows[i - 1]["MShare"]);

            //    tmpcash[4] += Convert.ToDouble(AlgFlow.Rows[i - 1]["ComCash"]);
            //    tmpshare[4] += Convert.ToInt32(AlgFlow.Rows[i - 1]["ComShare"]);

                
            //    tmprate[0] = ((tmpcash[0] + tmpshare[0]*Convert.ToDouble(a["price"])) / (Cash + Share*Convert.ToDouble(ALLSignals.Rows[0][2])))-1.0;

            //    tmprate[1] = ((tmpcash[1] + tmpshare[1] * Convert.ToDouble(a["price"])) / (Cash + Share * Convert.ToDouble(ALLSignals.Rows[0][2]))) - 1.0;

            //    tmprate[2] = ((tmpcash[2] + tmpshare[2] * Convert.ToDouble(a["price"])) / (Cash + Share * Convert.ToDouble(ALLSignals.Rows[0][2]))) - 1.0;

            //    tmprate[3] = ((tmpcash[3] + tmpshare[3] * Convert.ToDouble(a["price"])) / (Cash + Share * Convert.ToDouble(ALLSignals.Rows[0][2]))) - 1.0;

            //    tmprate[4] = ((tmpcash[4] + tmpshare[4] * Convert.ToDouble(a["price"])) / (Cash + Share * Convert.ToDouble(ALLSignals.Rows[0][2]))) - 1.0;

            //    tmprate[5] = (Convert.ToDouble(a["price"]) - Convert.ToDouble(ALLSignals.Rows[0][2])) * Share / (Cash + Share * Convert.ToDouble(ALLSignals.Rows[0][2]));

            //    AlgFlow.Rows.Add(a["DEven"], a["HEven"], a["price"], a["AlligatorFS"], tmpcash[0], tmpshare[0], tmprate[0],a["RSIFS"], tmpcash[1], tmpshare[1], tmprate[1],a["CCIFS"], tmpcash[2], tmpshare[2], tmprate[2],a["MACDFS"], tmpcash[3], tmpshare[3], tmprate[3],TotalSignals, tmpcash[4], tmpshare[4], tmprate[4],tmprate[5]);
                
            //    i++;
            //}
            //progressBar1.Value = 95;

           
            //dgSignalViewer.DataSource = AlgFlow;

            //progressBar1.Value = 100;

        }

        private void SetText(System.Windows.Forms.Control control, string text)
        {
            if (control.InvokeRequired)
            {
                SetTextCallback d = new SetTextCallback(SetText);
                Invoke(d, new object[] { control, text });
            }
            else
            {
                control.Text = text;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Optimizing();
            InitialCash = Convert.ToInt32(txtinitcash.Text);
            InitialShare = Convert.ToInt32(txtinitshare.Text);
            tooltxtstatus.Text = "Running...";
            toolProg.Visible = true;
            btnOptimize.Enabled = false;
            backgroundWorker1.RunWorkerAsync();
        }

        private void backgroundWorker1_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            Optimizing();
        }

        private void backgroundWorker1_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {

            toolProg.ProgressBar.Value = Math.Min(100, e.ProgressPercentage);
            //SetText(tooltxtstatus, e.UserState.ToString());
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {
            toolProg.Visible = false;
            tooltxtstatus.Text = "Done.";
            btnOptimize.Enabled = true;
        }

        private void Optimizing()
        {
            var d1 = DateTime.Now;
            string infoT  = string.Format("{0:####}{1:d2}{2:d2}_{3:##}{4:d2}{5:d2}.xls",
                                                   d1.Year,
                                                   d1.Month, d1.Day, d1.Hour, d1.Minute, d1.Second);
            

            string _outputFileNameT = "X:\\Output\\Total\\" + infoT;


            newFileT = new FileInfo(_outputFileNameT);
            xlPackageT = new ExcelPackage(newFileT);
            ExcelReport xlsRep = new ExcelReport();
            int NumberT=0;
            foreach (DataGridViewRow a in dgTaskGrid.Rows)
            {
                SetText(txtsymb, a.Cells["dgSymbol"].Value.ToString());
                SetText(txtstart, DateTime.Now.ToShortDateString());

                try
                {
                    ChromRange = new List<List<double>>();

                    Int32 AlgParam = Convert.ToInt32(a.Cells["dgParam"].Value);
                    Chromosome BestChrom = new Chromosome();
                    Scop = Convert.ToInt16(a.Cells["dgScope"].Value);
                    string Algorithm = a.Cells["dgAlg"].Value.ToString();
                    string Indicator = a.Cells["dgIndicator"].Value.ToString();


                    #region Indicator
                    switch (Indicator)
                    {
                        case "RSI":
                            range = new List<double>();
                            range.Add(Convert.ToDouble(txtRSI0L.Text));
                            range.Add(Convert.ToDouble(txtRSI0U.Text));
                            range.Add(Convert.ToDouble(chkRSI0.Checked));
                            range.Add(Convert.ToDouble(1));
                            ChromRange.Add(range);

                            range = new List<double>();
                            range.Add(Convert.ToDouble(txtRSI1L.Text));
                            range.Add(Convert.ToDouble(txtRSI1U.Text));
                            range.Add(Convert.ToDouble(chkRSI1.Checked));
                            range.Add(Convert.ToDouble(1));
                            ChromRange.Add(range);

                            range = new List<double>();
                            range.Add(Convert.ToDouble(txtRSI2L.Text));
                            range.Add(Convert.ToDouble(txtRSI2U.Text));
                            range.Add(Convert.ToDouble(chkRSI2.Checked));
                            range.Add(Convert.ToDouble(1));
                            ChromRange.Add(range);

                            range = new List<double>();
                            range.Add(Convert.ToDouble(txtRSI3L.Text));
                            range.Add(Convert.ToDouble(txtRSI3U.Text));
                            range.Add(Convert.ToDouble(chkRSI3.Checked));
                            range.Add(Convert.ToDouble(1));
                            ChromRange.Add(range);

                            break;

                        case "Alligator":
                            range = new List<double>();
                            range.Add(Convert.ToDouble(txtall0L.Text));
                            range.Add(Convert.ToDouble(txtall0U.Text));
                            range.Add(Convert.ToDouble(chkall0.Checked));
                            range.Add(Convert.ToDouble(txtAll0S.Text));
                            ChromRange.Add(range);

                            range = new List<double>();
                            range.Add(Convert.ToDouble(txtall1L.Text));
                            range.Add(Convert.ToDouble(txtall1U.Text));
                            range.Add(Convert.ToDouble(chkall1.Checked));
                            range.Add(Convert.ToDouble(txtAll1S.Text));
                            ChromRange.Add(range);

                            range = new List<double>();
                            range.Add(Convert.ToDouble(txtall2L.Text));
                            range.Add(Convert.ToDouble(txtall2U.Text));
                            range.Add(Convert.ToDouble(chkall2.Checked));
                            range.Add(Convert.ToDouble(txtAll2S.Text));
                            ChromRange.Add(range);

                            range = new List<double>();
                            range.Add(Convert.ToDouble(txtall3L.Text));
                            range.Add(Convert.ToDouble(txtall3U.Text));
                            range.Add(Convert.ToDouble(chkall3.Checked));
                            range.Add(Convert.ToDouble(txtAll3S.Text));
                            ChromRange.Add(range);

                            range = new List<double>();
                            range.Add(Convert.ToDouble(txtall4L.Text));
                            range.Add(Convert.ToDouble(txtall4U.Text));
                            range.Add(Convert.ToDouble(chkall4.Checked));
                            range.Add(Convert.ToDouble(txtAll4S.Text));
                            ChromRange.Add(range);

                            range = new List<double>();
                            range.Add(Convert.ToDouble(txtall5L.Text));
                            range.Add(Convert.ToDouble(txtall5U.Text));
                            range.Add(Convert.ToDouble(chkall5.Checked));
                            range.Add(Convert.ToDouble(txtAll5S.Text));
                            ChromRange.Add(range);

                            range = new List<double>();
                            range.Add(Convert.ToDouble(txtall6L.Text));
                            range.Add(Convert.ToDouble(txtall6U.Text));
                            range.Add(Convert.ToDouble(chkall6.Checked));
                            range.Add(Convert.ToDouble(txtAll6S.Text));
                            ChromRange.Add(range);

                            range = new List<double>();
                            range.Add(Convert.ToDouble(txtall7L.Text));
                            range.Add(Convert.ToDouble(txtall7U.Text));
                            range.Add(Convert.ToDouble(chkall7.Checked));
                            range.Add(Convert.ToDouble(txtall7S.Text));
                            ChromRange.Add(range);
                            
                            break;

                        case "MACD":
                            range = new List<double>();
                            range.Add(Convert.ToDouble(txtmacd0l.Text));
                            range.Add(Convert.ToDouble(txtmacd0u.Text));
                            range.Add(Convert.ToDouble(chkmacd0.Checked));
                            range.Add(Convert.ToDouble(1));
                            ChromRange.Add(range);

                            range = new List<double>();
                            range.Add(Convert.ToDouble(txtmacd1l.Text));
                            range.Add(Convert.ToDouble(txtmacd1u.Text));
                            range.Add(Convert.ToDouble(chkmacd1.Checked));
                            range.Add(Convert.ToDouble(1));
                            ChromRange.Add(range);

                            range = new List<double>();
                            range.Add(Convert.ToDouble(txtmacd2l.Text));
                            range.Add(Convert.ToDouble(txtmacd2u.Text));
                            range.Add(Convert.ToDouble(chkmacd2.Checked));
                            range.Add(Convert.ToDouble(1));
                            ChromRange.Add(range);

                            range = new List<double>();
                            range.Add(Convert.ToDouble(txtmacd3l.Text));
                            range.Add(Convert.ToDouble(txtmacd3u.Text));
                            range.Add(Convert.ToDouble(chkmacd3.Checked));
                            range.Add(Convert.ToDouble(1));
                            ChromRange.Add(range);

                            break;

                        case "CCI":
                            range = new List<double>();
                            range.Add(Convert.ToDouble(txtCCI0L.Text));
                            range.Add(Convert.ToDouble(txtCCI0U.Text));
                            range.Add(Convert.ToDouble(chkCCI0.Checked));
                            range.Add(Convert.ToDouble(1));
                            ChromRange.Add(range);

                            range = new List<double>();
                            range.Add(Convert.ToDouble(txtCCI1L.Text));
                            range.Add(Convert.ToDouble(txtCCI1U.Text));
                            range.Add(Convert.ToDouble(chkCCI1.Checked));
                            range.Add(Convert.ToDouble(1));
                            ChromRange.Add(range);

                            range = new List<double>();
                            range.Add(Convert.ToDouble(txtCCI2L.Text));
                            range.Add(Convert.ToDouble(txtCCI2U.Text));
                            range.Add(Convert.ToDouble(chkCCI2.Checked));
                            range.Add(Convert.ToDouble(1));
                            ChromRange.Add(range);

                            range = new List<double>();
                            range.Add(Convert.ToDouble(txtCCI3L.Text));
                            range.Add(Convert.ToDouble(txtCCI3U.Text));
                            range.Add(Convert.ToDouble(chkCCI3.Checked));
                            range.Add(Convert.ToDouble(1));
                            ChromRange.Add(range);

                            break;
                    }
                    #endregion

                    
                    var d = DateTime.Now;
                    string info = a.Cells["dgInsCode"].Value + "_" + Indicator + "_" + AlgParam;
                    OutputFileName = string.Format(info + "_{0:####}{1:d2}{2:d2}_{3:##}{4:d2}{5:d2}.xls",
                                                   d.Year,
                                                   d.Month, d.Day, d.Hour, d.Minute, d.Second);

                    string _outputFileName = a.Cells["dgDest"].Value.ToString() + "/" + OutputFileName;


                    newFile = new FileInfo(_outputFileName);
                    xlPackage = new ExcelPackage(newFile);

                    Optimization OptFunc = new Optimization();
                    OptFunc.SDEven = Convert.ToInt32(a.Cells["dgSDeven"].Value);
                    OptFunc.SHEven = Convert.ToInt32(a.Cells["dgSHeven"].Value);
                    OptFunc.EDEven = Convert.ToInt32(a.Cells["dgEDeven"].Value);
                    OptFunc.EHEven = Convert.ToInt32(a.Cells["dgEHeven"].Value);
                    OptFunc.Scope = Convert.ToInt32(a.Cells["dgScope"].Value);
                    OptFunc.Indicator = Indicator;
                    OptFunc.xlPackage = xlPackage;
                    OptFunc.ChromRange = ChromRange;
                    OptFunc.Cash = Convert.ToInt64(InitialCash);
                    OptFunc.Share = Convert.ToInt64(InitialShare);
                    OptFunc.mySql = mySql;
                    OptFunc.Algorithm = Algorithm;

                    a.Cells["dgStatus"].Value = "In Progress...";

                    OptFunc.Symb = new Symbols(a.Cells["dgSymbol"].Value.ToString(), a.Cells["dgInsCode"].Value.ToString());

                    switch (Algorithm)
                    {
                        case "Genetic Algorithm":
                            BestChrom = OptFunc.GeneticAlg(AlgParam);
                            break;
                        case "Grid Search":
                            BestChrom = OptFunc.GridSearch(AlgParam);
                            break;
                        case "Greedy Grid":
                            BestChrom = OptFunc.GreedyGrid(AlgParam);
                            break;
                    }

                    a.Cells["dgStatus"].Value = "Done-" + DateTime.Now.ToShortDateString();

                    xlsRep.ExcelReport_ByChrom(BestChrom, OptFunc, xlPackageT, NumberT, "Overal Results");
                    
                    xlPackage.Workbook.Properties.Title = "Indicator Optimization, ver 2.6.1, July 2013";
                    xlPackage.Workbook.Properties.Author = "Mohsen Dastpak";
                    xlPackage.Workbook.Properties.Manager = "Mohsen Dastpak";
                    xlPackage.Workbook.Properties.Company = "Vtoc";
                    xlPackage.Workbook.Properties.SetCustomPropertyValue("EmployeeID", "2076");
                    xlPackage.Save();

                }
                catch(Exception e)
                {
                    a.Cells["dgStatus"].Value = e.Message + "_" + DateTime.Now.ToShortDateString();
                }
                NumberT++;
            }

            xlPackageT.Workbook.Properties.Title = "Indicator Optimization, ver 2.6.1, July 2013";
            xlPackageT.Workbook.Properties.Author = "Mohsen Dastpak";
            xlPackageT.Workbook.Properties.Manager = "Mohsen Dastpak";
            xlPackageT.Workbook.Properties.Company = "Vtoc";
            xlPackageT.Workbook.Properties.SetCustomPropertyValue("EmployeeID", "2076");
            xlPackageT.Save();

        }

        private void lstFirstSymbol_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            btnoneright_Click(sender, e);
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            int i = 1;

            foreach (var a in FinalSymbol)
            {
                string Algorithm = "Standard";
                var d = DateTime.Now;
                string info = a.InsCode + "_Standard_" + txtalgparam.Text;
                OutputFileName = string.Format(info + "_{0:####}{1:d2}{2:d2}_{3:##}{4:d2}{5:d2}.xls",
                                               d.Year,
                                               d.Month, d.Day, d.Hour, d.Minute, d.Second);

                string _outputFileName = txtaddress.Text + "/" + OutputFileName;
                Scop = Convert.ToInt16(cmbScop.Text);

                newFile = new FileInfo(_outputFileName);
                xlPackage = new ExcelPackage(newFile);

            
                ProfitFunction PF = new ProfitFunction();
                PF.SDEven = Convert.ToInt32(txtSDEven.Text);
                PF.SHEven = Convert.ToInt32(txtSHEven.Text);
                PF.EDEven = Convert.ToInt32(txtDEven.Text);
                PF.EHEven = Convert.ToInt32(txtHEven.Text);
                PF.Scope = Scop;
                PF.Indicator = cmbalg.Text;
                PF.Cash = Convert.ToInt64(txtinitcash.Text);
                PF.Share = Convert.ToInt64(txtinitshare.Text);
                PF.mySql = mySql;
                PF.InsCode = a.InsCode;
                PF.dataCache = new Dictionary<string, DataTable>();

                List<double> Genes = new List<double>();
                switch (cmbalg.Text)
                {
                    case "Alligator":
                        Genes.Add(3);
                        Genes.Add(13);
                        Genes.Add(8);
                        Genes.Add(5);
                        Genes.Add(8);
                        Genes.Add(5);
                        Genes.Add(3);
                        Genes.Add(0.2);
                        
                        PF.Genes = Genes;
                        break;
                }
                Chromosome tmpChr = new Chromosome();
                tmpChr.Genes = Genes;
                tmpChr.J = PF.ProfitFunctionT();
                tmpChr.SignalsNum = PF.SignalsNum;

                List<Chromosome> tmplstchr = new List<Chromosome>();
                tmplstchr.Add(tmpChr);

                ExcelReport XLS = new ExcelReport();

                XLS.ExcelReportT_Overal(tmplstchr, i, 1, xlPackage, i, a.Symbol + "_Standard");
                XLS.ExcelReport_General(PF.Buy_Hold().ToString(), xlPackage, 1, "B_H");


                xlPackage.Workbook.Properties.Title = "Indicator Optimization, ver 2.6.1, July 2013";
                xlPackage.Workbook.Properties.Author = "Mohsen Dastpak";
                xlPackage.Workbook.Properties.Manager = "Mohsen Dastpak";
                xlPackage.Workbook.Properties.Company = "Vtoc";
                xlPackage.Workbook.Properties.SetCustomPropertyValue("EmployeeID", "2076");
                xlPackage.Save();

                i++;
            }
            
            

        }

        private void button2_Click(object sender, EventArgs e)
        {
            List<double> Genes = new List<double>();
            int i = 1;
            foreach (var a in FinalSymbol)
            {
                string Algorithm = "Specific";
                var d = DateTime.Now;
                string info = a.InsCode + "_Specific_" + txtalgparam.Text;
                OutputFileName = string.Format(info + "_{0:####}{1:d2}{2:d2}_{3:##}{4:d2}{5:d2}.xls",
                                               d.Year,
                                               d.Month, d.Day, d.Hour, d.Minute, d.Second);

                string _outputFileName = txtaddress.Text + "/" + OutputFileName;
                Scop = Convert.ToInt16(cmbScop.Text);

                newFile = new FileInfo(_outputFileName);
                xlPackage = new ExcelPackage(newFile);

                ProfitFunction PF = new ProfitFunction();
                PF.SDEven = Convert.ToInt32(txtSDEven.Text);
                PF.SHEven = Convert.ToInt32(txtSHEven.Text);
                PF.EDEven = Convert.ToInt32(txtDEven.Text);
                PF.EHEven = Convert.ToInt32(txtHEven.Text);
                PF.Scope = Scop;
                PF.Indicator = cmbalg.Text;
                PF.Cash = Convert.ToInt64(txtinitcash.Text);
                PF.Share = Convert.ToInt64(txtinitshare.Text);
                PF.mySql = mySql;
                PF.InsCode = a.InsCode;
                PF.dataCache = new Dictionary<string, DataTable>();
                Genes = new List<double>();              
                switch (cmbalg.Text)
                {
                    case "Alligator":
                        Genes.Add(Convert.ToDouble(txtall0O.Text));
                        Genes.Add(Convert.ToDouble(txtall1O.Text));
                        Genes.Add(Convert.ToDouble(txtall2O.Text));
                        Genes.Add(Convert.ToDouble(txtall3O.Text));
                        Genes.Add(Convert.ToDouble(txtall4O.Text));
                        Genes.Add(Convert.ToDouble(txtall5O.Text));
                        Genes.Add(Convert.ToDouble(txtall6O.Text));
                        Genes.Add(Convert.ToDouble(txtall7O.Text));

                        PF.Genes = Genes;
                        break;
                    case "RSI":
                        Genes.Add(Convert.ToDouble(txtRSI0O.Text));
                        Genes.Add(Convert.ToDouble(txtRSI1O.Text));
                        Genes.Add(Convert.ToDouble(txtRSI2O.Text));
                        Genes.Add(Convert.ToDouble(txtRSI3O.Text));
                        break;
                }
                Chromosome tmpChr = new Chromosome();
                tmpChr.Genes = Genes;
                PF.Genes = Genes;
                tmpChr.J = PF.ProfitFunctionT();
                tmpChr.SignalsNum = PF.SignalsNum;

                List<Chromosome> tmplstchr = new List<Chromosome>();
                tmplstchr.Add(tmpChr);

                ExcelReport XLS = new ExcelReport();

                XLS.ExcelReportT_Overal(tmplstchr, 1, 1, xlPackage, i, a.Symbol + "_Standard");

                xlPackage.Workbook.Properties.Title = "Indicator Optimization, ver 2.6.1, July 2013";
                xlPackage.Workbook.Properties.Author = "Mohsen Dastpak";
                xlPackage.Workbook.Properties.Manager = "Mohsen Dastpak";
                xlPackage.Workbook.Properties.Company = "Vtoc";
                xlPackage.Workbook.Properties.SetCustomPropertyValue("EmployeeID", "2076");
                xlPackage.Save();

                i++;
            }

        }

        private void btnaddtask_Click(object sender, EventArgs e)
        {
            foreach (var a in FinalSymbol)
            {   
                dgTaskGrid.Rows.Add("Del",DateTime.Now.ToShortDateString(), a.Symbol, a.InsCode, txtSDEven.Text, txtSHEven.Text, txtDEven.Text, txtHEven.Text, cmbScop.Text, cmbalg.Text, cmbOptAlg.Text, txtalgparam.Text, txtaddress.Text); 
                
            }
        }

        
        private void btnimport_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
            string txtfile = "";
            if (openFileDialog1.FileName != "")
            {
                txtfile = System.IO.File.ReadAllText(openFileDialog1.FileName);
                lstFinalSymbol.Items.Clear();
                FinalSymbol.Clear();

                string[] tmptxt1 = txtfile.Split('$');
                foreach (var a in tmptxt1)
                {
                    string[] tmptxt2 = a.Split('!');

                    Symbols tmpsymb = new Symbols(tmptxt2[0], tmptxt2[1]);
                    lstFinalSymbol.Items.Add(tmpsymb);
                    lstFinalSymbol.ValueMember = "InsCode";
                    lstFinalSymbol.DisplayMember = "Symbol";
                    lstFinalSymbol.SelectedIndex = lstFinalSymbol.Items.Count - 1;

                    FinalSymbol.Add(tmpsymb);
                }
            }
        }

        private void btnexport_Click(object sender, EventArgs e)
        {
            string txtfile = "";
            saveFileDialog1.ShowDialog();
            if (saveFileDialog1.FileName != "")
            {
                foreach (Symbols a in FinalSymbol)
                {
                    txtfile += a.Symbol + "!" + a.InsCode + "$";
                }
                txtfile = txtfile.Substring(0, txtfile.Count() - 1);

                System.IO.File.WriteAllText(saveFileDialog1.FileName, txtfile);

            }
        }
        
    }
}




public class Chromosome
{
    public List<double> Genes;
    public double J { get; set; }
    public double RW_N;
    public double RW_Acc;
    public int SignalsNum;
    public Chromosome()
    {
        Genes = new List<double>();
        J = 0;

    }

    public Chromosome(List<double> mygene,double myJ, int mySignals)
    {
        Genes = mygene;
        J = myJ;
        RW_N = J + 5;
        SignalsNum = mySignals;
    }


}

