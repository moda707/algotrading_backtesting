using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using vtocSqlInterface;
using System.Data;
using AForge;
using AForge.Neuro;
using AForge.Neuro.Learning;


using Accord.Neuro;
using Accord.Neuro.Learning;
using AForge.Controls;

namespace BackTesting
{
    public class ProfitFunction
    {
        public string InsCode { get; set; }
        public int SDEven { get; set; }
        public int SHEven { get; set; }
        public int EDEven { get; set; }
        public int EHEven { get; set; }
        public string Indicator { get; set; }
        public int Scope { get; set; }
        public int MPeriod { get; set; }
        public int Step { get; set; }
        public List<double> Genes { get; set; }
        public int SignalsNum { get; set; }
        public Dictionary<string, DataTable> dataCache { get; set; }
        
        public ActivationNetwork myNetwork { get; set; }

        public Int64 Cash { get; set; }
        public Int64 Share { get; set; }


        public vtocSqlInterface.sqlInterface mySql { get; set; }

        private string sqlCmd;

        public double Calculating(DataTable ALLSignals)
        {
            double RetProfit = 0;
            
            List<double> AlgFlow = new List<double>();

            AlgFlow.Add(0);//0-price
            AlgFlow.Add(Cash);//1-Cash
            AlgFlow.Add(Share);//2-Share
            AlgFlow.Add(0);//3-Rate


            int i = 0;
            SignalsNum = 0;
            foreach (DataRow a in ALLSignals.Rows)
            {
                i++;
                double tmpcash = 0;
                double tmpshare = 0;
                double tmprate = 0;
                //Alligator
                if (Convert.ToInt32(a["Signal"]) > 0 && a["QTitMeOf"].ToString() != "")
                {
                    tmpshare = Math.Max(Math.Min((Int32)(0.2 * Convert.ToDouble(AlgFlow[1]) / Convert.ToDouble(a["PMeOf"])), Convert.ToInt64(a["QTitMeOf"])), 0);
                    tmpcash = -1.005 * Convert.ToDouble(a["PMeOf"]) * tmpshare;
                    SignalsNum++;
                }
                else if (Convert.ToInt32(a["Signal"]) < 0 && a["QTitMeDem"].ToString() != "")
                {
                    tmpshare = -Math.Max(Math.Min(Convert.ToInt32(AlgFlow[2]), Convert.ToInt64(a["QTitMeDem"])), 0);
                    tmpcash = 0.99 * (-tmpshare) * Convert.ToDouble(a["PMeDem"]);
                    SignalsNum++;
                }
                tmpcash += Convert.ToDouble(AlgFlow[1]);
                tmpshare += Convert.ToInt32(AlgFlow[2]);
                tmprate = ((tmpcash + tmpshare * Convert.ToDouble(a["price"])) / (Cash + Share * Convert.ToDouble(ALLSignals.Rows[0][2]))) - 1.0;

                AlgFlow[0] = Convert.ToDouble(a["price"]);
                AlgFlow[1] = tmpcash;
                AlgFlow[2] = tmpshare;
                AlgFlow[3] = tmprate;

                RetProfit = tmprate;
            }

            return RetProfit;
        }

        public double ProfitFunctionT()
        {
            
            mySql = new sqlInterface(Properties.Settings.Default.sqlserver, "TseTrade",
                                     Properties.Settings.Default.username, Properties.Settings.Default.pass);
            int NHC = Genes.Count;

            sqlCmd = @"SELECT T.DEven, T.HEven, T.price, ((CASE WHEN T.Buy>0 THEN 1 ELSE 0 END) + (CASE WHEN T.Sell>0 THEN -1 ELSE 0 END)) Signal, T.PMeDem, T.PMeOf, T.QTitMeDem, T.QTitMeOf FROM TseTrade.dbo.fn_AT_IND_" + Indicator + "_SIGNALS(" + InsCode + "," + Scope + "," + SDEven + "," + SHEven + "," + EDEven + "," + EHEven;

            foreach (var a in Genes)
            {
                sqlCmd += "," + a.ToString();
            }
            sqlCmd += ") T ORDER BY T.DEven,T.HEven";

            DataTable ALLSignals;
            if (dataCache.ContainsKey(sqlCmd))
                ALLSignals = dataCache[sqlCmd];
            else
            {
                ALLSignals = mySql.SqlExecuteReader(sqlCmd);
                dataCache.Add(sqlCmd, ALLSignals);
            }

            return Calculating(ALLSignals);
        }

        public double Buy_Hold()
        {
            mySql = new sqlInterface(Properties.Settings.Default.sqlserver, "TseTrade",
                                     Properties.Settings.Default.username, Properties.Settings.Default.pass);

            sqlCmd = "SELECT TOP (1) W.PriceAvg FROM TseTrade.dbo.TsePriceCandles W WHERE W.InsCode = " + InsCode + " AND W.Periodicity = " + Scope + "  AND ((W.DEven=" + EDEven + " AND W.HEven<=" + EHEven + ") OR (W.DEven<" + EDEven + ")) AND ((W.DEven=" + SDEven + " AND W.HEven>=" + SHEven + ") OR (W.DEven>" + SDEven + ")) ORDER BY W.DEven, W.HEven";
            DataTable FPrice = new DataTable();
            FPrice = mySql.SqlExecuteReader(sqlCmd);

            sqlCmd = "SELECT TOP (1) W.PriceAvg FROM TseTrade.dbo.TsePriceCandles W WHERE W.InsCode = " + InsCode + " AND W.Periodicity = " + Scope + "  AND ((W.DEven=" + EDEven + " AND W.HEven<=" + EHEven + ") OR (W.DEven<" + EDEven + ")) AND ((W.DEven=" + SDEven + " AND W.HEven>=" + SHEven + ") OR (W.DEven>" + SDEven + ")) ORDER BY W.DEven DESC, W.HEven DESC";
            DataTable LPrice = new DataTable();
            LPrice = mySql.SqlExecuteReader(sqlCmd);


            double FirstPrice = Convert.ToDouble(FPrice.Rows[0][0]);
            double LastPrice = Convert.ToDouble(LPrice.Rows[0][0]);

            return ((LastPrice - FirstPrice) / FirstPrice);
        }

        public double Neural_Test()
        {            
            int InputLayerCount = 12;
            mySql = new sqlInterface(Properties.Settings.Default.sqlserver, "TseTrade",
                                     Properties.Settings.Default.username, Properties.Settings.Default.pass);

            sqlCmd = "SELECT * FROM dbo.fn_AT_IND_NeuralBatch(" + InsCode + "," + Scope + "," + SDEven + "," + SHEven + ", " + EDEven + ", " + EHEven + ", " + MPeriod + "," + Step + ")";

            DataTable ALLSignals;
            ALLSignals = mySql.SqlExecuteReader(sqlCmd);

            ALLSignals.Columns.Add("Signal");

            foreach (DataRow a in ALLSignals.Rows)
            {

                double[] input = new double[InputLayerCount];
                for (int i = 0; i < InputLayerCount; i++)
                {
                    input[i] = Convert.ToDouble(a[8 + i]);
                }

                double [] output = myNetwork.Compute(input);

                if (output[0] > 0.5 && output[1] <= 0.5)
                {
                    a["Signal"] = -1;
                }
                else if (output[1] > 0.5 && output[0] <= 0.5)
                {
                    a["Signal"] = 1;
                }
                else
                {
                    a["Signal"] = 0;
                }
            }


            return Calculating(ALLSignals);
        }


    }
}
