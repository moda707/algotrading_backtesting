using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using System.Data;

namespace BackTesting
{
    class Optimization
    {
        public Random RVar = new Random();
        public Symbols Symb { get; set; }
        public List<List<double>> ChromRange { get; set; }
        public string Indicator {get;set;}
        public ExcelPackage xlPackage {get;set;}
        public int SDEven { get; set; }
        public int SHEven { get; set; }
        public int EDEven { get; set; }
        public int EHEven { get; set; }        
        public int Scope { get; set; }
        public Int64 Cash {get;set;}
        public Int64 Share {get;set;}
        public int Generation { get; set; }
        public vtocSqlInterface.sqlInterface mySql;
        public string Algorithm { get; set; }

        private ProfitFunction PF = new ProfitFunction();
        private ExcelReport XLS = new ExcelReport();

        public Chromosome GeneticAlg(int Population)
        {            
            int NCH = ChromRange.Count;
            var ChromListPool = new List<Chromosome>();
            ProfitFunction PF = new ProfitFunction();
            Chromosome OptChrom = new Chromosome();

            //backgroundWorker1.ReportProgress(1, "Populating");
            double TProfit;
            int SignalsNum;

            //Make a pool
            int i = 0;
            while (i < 5 * Population)
            {
                try
                {
                    var t = new List<double>();
                    for (int j = 0; j < NCH; j++)
                    {
                        t.Add(GetRandom(ChromRange[j]));
                    }

                    PF = new ProfitFunction();
                    PF.Indicator = Indicator;
                    PF.InsCode = Symb.InsCode;
                    PF.SDEven = SDEven;
                    PF.SHEven = SHEven;
                    PF.EDEven = EDEven;
                    PF.EHEven = EHEven;
                    PF.Genes = t;
                    PF.Scope = Scope;
                    PF.dataCache = dataCache;
                    PF.Cash = Cash;
                    PF.Share = Share;
                    PF.mySql = mySql;
                    TProfit = PF.ProfitFunctionT();

                    SignalsNum = PF.SignalsNum;

                    if (SignalsNum > 0)
                    {
                        ChromListPool.Add(new Chromosome(t, TProfit, SignalsNum));
                        i++;
                    }
                }
                catch (Exception e)
                {
                    XLS.ExcelReport_Log(e.Message, xlPackage, i, Symb.Symbol + "_Init");
                }
            }


            
            //First Selection from Pool
            var ChromList = new List<Chromosome>();
            //ChromList = ChromListPool.GetRange(0, Population);
            for (int i3 = 0; i3 < Population; i3++)
            {
                ChromList.Add(ChromListPool[RVar.Next(4 * Population)]);
            }

            ChromList = ChromList.OrderByDescending(s => s.J).ToList();

            #region poolExcel
            
            XLS.ExcelReportT_Overal(ChromListPool, 1, 1, xlPackage, 0, Symb.Symbol);
            List<Chromosome> tmpPool = new List<Chromosome>();
            tmpPool = ChromListPool.GetRange(0, 1);
            XLS.ExcelReportT_Overal(tmpPool, 1, 1, xlPackage, 0, Symb.Symbol + "_Summary");                        
            #endregion


            //Generations
            int notchanged = 0;
            int Gener = 0;
            
            while (notchanged <= 6)            
            {
                var tmpChromList = new List<Chromosome>();
                tmpChromList = ChromList;

                //GENERATE THE POPULATION
                //Elitism : Get 10% of Top of each Chrom list
                int ElitNum = (int)(0.1 * Population);

                int j = ElitNum;
                while (j < Population)
                {

                    double RandNum;

                    Chromosome Parent1 = FindParent(tmpChromList);
                    List<Chromosome> tmpch = new List<Chromosome>();
                    tmpch.Add(Parent1);
                    tmpch = tmpChromList.Except(tmpch).ToList();
                    Chromosome Parent2 = FindParent(tmpch);


                    Chromosome Child1 = new Chromosome();
                    Chromosome Child2 = new Chromosome();

                    int GeneInd1 = RVar.Next(NCH);

                    var Gene1 = new List<double>();
                    var Gene2 = new List<double>();

                    for (int p = 0; p <= GeneInd1; p++)
                    {
                        Gene1.Add(Parent1.Genes[p]);
                        Gene2.Add(Parent2.Genes[p]);
                    }
                    for (int p = GeneInd1 + 1; p < NCH; p++)
                    {
                        Gene1.Add(Parent2.Genes[p]);
                        Gene2.Add(Parent1.Genes[p]);
                    }



                    //Mutation
                    for (int w = 0; w < Gene1.Count; w++)
                    {
                        RandNum = RVar.NextDouble();
                        if (RandNum < 0.09)
                        {
                            Gene1[w] = GetRandom(ChromRange[w]);
                        }

                        RandNum = RVar.NextDouble();
                        if (RandNum < 0.09)
                        {
                            Gene2[w] = GetRandom(ChromRange[w]);
                        }
                    }

                    try
                    {
                        PF = new ProfitFunction();
                        PF.Indicator = Indicator;
                        PF.InsCode = Symb.InsCode;
                        PF.SDEven = SDEven;
                        PF.SHEven = SHEven;
                        PF.EDEven = EDEven;
                        PF.EHEven = EHEven;
                        PF.Genes = Gene1;
                        PF.Scope = Scope;
                        PF.dataCache = dataCache;
                        PF.Cash = Cash;
                        PF.Share = Share;

                        TProfit = PF.ProfitFunctionT();

                        SignalsNum = PF.SignalsNum;

                        Child1 = new Chromosome(Gene1, TProfit, SignalsNum);

                        PF = new ProfitFunction();
                        PF.Indicator = Indicator;
                        PF.InsCode = Symb.InsCode;
                        PF.SDEven = SDEven;
                        PF.SHEven = SHEven;
                        PF.EDEven = EDEven;
                        PF.EHEven = EHEven;
                        PF.Genes = Gene2;
                        PF.Scope = Scope;
                        PF.dataCache = dataCache;
                        PF.Cash = Cash;
                        PF.Share = Share;

                        TProfit = PF.ProfitFunctionT();

                        SignalsNum = PF.SignalsNum;
                        Child2 = new Chromosome(Gene2, TProfit, SignalsNum);


                        if (Child1.SignalsNum > 0 && Child2.SignalsNum > 0)
                        {
                            if (Child1.J > Child2.J)
                            {
                                ChromList[j] = Child1;
                            }
                            else
                            {
                                ChromList[j] = Child2;
                            }
                            j++;
                        }
                        else if (Child1.SignalsNum > 0 && Child2.SignalsNum == 0)
                        {
                            ChromList[j] = Child1;
                            j++;
                        }
                        else if (Child2.SignalsNum > 0 && Child1.SignalsNum == 0)
                        {
                            ChromList[j] = Child2;
                            j++;
                        }
                    }
                    catch (Exception e)
                    {
                        XLS.ExcelReport_Log(e.Message, xlPackage, Gener, Symb.Symbol);
                    }
                }

                ChromList = ChromList.OrderByDescending(s => s.J).ToList();

                int tStartR = 2 + (Gener) * (Population + 1) + ChromListPool.Count;
                int tStartC = 1;

                XLS.ExcelReportT_Overal(ChromList, tStartR, tStartC, xlPackage, Gener, Symb.Symbol);
                List<Chromosome> tmpchr = new List<Chromosome>();
                tmpchr.Add(ChromList[0]);
                XLS.ExcelReportT_Overal(tmpchr, Gener + 1, 1, xlPackage, Gener, Symb.Symbol + "_Summary");

                if (ChromList[0].J == tmpChromList[0].J) notchanged++; else notchanged = 0; //Criteria for Algorithm stopage

                Gener++;
            }

            Generation = Gener;
            OptChrom = ChromList[0];

            return OptChrom;
        }

        public Chromosome FindParent(List<Chromosome> ChrmLst)
        {
            Chromosome Parent = new Chromosome();
            double TotalWeight = ChrmLst.Sum(t => t.RW_N);

            double PrVal = 0;
            var RandNum = 0.0;
            foreach (var w in ChrmLst)
            {
                w.RW_Acc = PrVal + w.RW_N / TotalWeight;
                PrVal = w.RW_Acc;
            }

            //Select Parent
            RandNum = RVar.NextDouble();
            Parent = ChrmLst.Find(t => (t.RW_Acc >= RandNum && RandNum > (t.RW_Acc - (t.RW_N / TotalWeight))));

            if (Parent == null) Parent = ChrmLst[RVar.Next(ChrmLst.Count)];

            return Parent;
        }

        public double GetRandom(List<double> Range)
        {
            double RNumber = 0.0;
            double Lower = Range[0];
            double Upper = Range[1];
            double IsInt = Range[2];
            int Step = (int)Range[3];


            double tmprnd = RVar.NextDouble();
            if ((int)IsInt == 1)
            {
                RNumber = Lower + Math.Round(((Upper - Lower) / Step) * (tmprnd)) * Step;
            }
            else
            {
                RNumber = Lower + (Upper - Lower) * (tmprnd);
            }

            return RNumber;
        }

        Dictionary<string, DataTable> dataCache = new Dictionary<string, DataTable>();

        //Grid Search
        public Chromosome GridSearch(int GridNum)
        {
            Chromosome BestChrom = new Chromosome();
            List<double> Gens = new List<double>();
            int Step = GridNum;

            BestChrom.Genes.Add(0);
            BestChrom.Genes.Add(0);
            BestChrom.Genes.Add(0);
            BestChrom.Genes.Add(0);
            BestChrom.Genes.Add(0);
            BestChrom.Genes.Add(0);
            BestChrom.Genes.Add(0);
            BestChrom.Genes.Add(0);

            Generation = 0;
                        
            //backgroundWorker1.ReportProgress((stock - 1) * (100 / lstFinalSymbol.Items.Count), "Start Searching (" + stock + "/" + lstFinalSymbol.Items.Count + ")");
            switch (Indicator)
            {
                case "Alligator":
                    //for Alligator with 7 Parameter
                    for (int p1 = (int)ChromRange[0][0]; p1 <= (int)ChromRange[0][1]; p1 += Step)
                    {
                        for (int p2 = (int)ChromRange[1][0]; p2 <= (int)ChromRange[1][1]; p2 += Step)
                        {
                            for (int p3 = (int)ChromRange[2][0]; p3 <= (int)ChromRange[2][1]; p3 += Step)
                            {
                                for (int p4 = (int)ChromRange[3][0]; p4 <= (int)ChromRange[3][1]; p4 += Step)
                                {
                                    for (int p5 = (int)ChromRange[4][0]; p5 <= (int)ChromRange[4][1]; p5 += Step)
                                    {
                                        for (int p6 = (int)ChromRange[5][0]; p6 <= (int)ChromRange[5][1]; p6 += Step)
                                        {
                                            for (int p7 = (int)ChromRange[6][0]; p7 <= (int)ChromRange[6][1]; p7 += Step)
                                            {
                                                int SignalsNum = 0;
                                                Gens = new List<double>();
                                                Gens.Add(p1);
                                                Gens.Add(p2);
                                                Gens.Add(p3);
                                                Gens.Add(p4);
                                                Gens.Add(p5);
                                                Gens.Add(p6);
                                                Gens.Add(p7);
                                                Gens.Add(GetRandom(ChromRange[7]));
                                                try
                                                {
                                                    double tmpProfit;
                                                    PF = new ProfitFunction();
                                                    PF.Indicator = Indicator;
                                                    PF.InsCode = Symb.InsCode;
                                                    PF.SDEven = SDEven;
                                                    PF.SHEven = SHEven;
                                                    PF.EDEven = EDEven;
                                                    PF.EHEven = EHEven;
                                                    PF.Genes = Gens;
                                                    PF.Scope = Scope;
                                                    PF.dataCache = dataCache;

                                                    tmpProfit = PF.ProfitFunctionT();

                                                    SignalsNum = PF.SignalsNum;

                                                    if (tmpProfit > BestChrom.J || Generation == 0)
                                                    {
                                                        BestChrom.J = tmpProfit;
                                                        BestChrom.Genes = Gens;
                                                        BestChrom.SignalsNum = SignalsNum;
                                                    }
                                                    List<Chromosome> tmpList = new List<Chromosome>();
                                                    tmpList.Add(BestChrom);
                                                    XLS.ExcelReportT_Overal(tmpList, Generation + 1, 1, xlPackage, Generation, Symb.Symbol);

                                                }
                                                catch (Exception e)
                                                {
                                                    XLS.ExcelReport_Log(e.Message, xlPackage, Generation, Symb.Symbol);                                                    
                                                }

                                                Generation++;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    break;
                case "RSI":
                    //for RSI with 4 Parameter
                    for (int p1 = (int)ChromRange[0][0]; p1 <= (int)ChromRange[0][1]; p1 += Step)
                    {
                        for (int p2 = (int)ChromRange[1][0]; p2 <= (int)ChromRange[1][1]; p2 += Step)
                        {
                            for (int p3 = (int)ChromRange[2][0]; p3 <= (int)ChromRange[2][1]; p3 += Step)
                            {
                                for (int p4 = (int)ChromRange[3][0]; p4 <= (int)ChromRange[3][1]; p4 += Step)
                                {

                                    int SignalsNum = 0;
                                    Gens = new List<double>();
                                    Gens.Add(p1);
                                    Gens.Add(p2);
                                    Gens.Add(p3);
                                    Gens.Add(p4);
                                    Gens.Add(GetRandom(ChromRange[4]));


                                    try
                                    {
                                        double tmpProfit;
                                        PF = new ProfitFunction();
                                        PF.Indicator = Indicator;
                                        PF.InsCode = Symb.InsCode;
                                        PF.SDEven = SDEven;
                                        PF.SHEven = SHEven;
                                        PF.EDEven = EDEven;
                                        PF.EHEven = EHEven;
                                        PF.Genes = Gens;
                                        PF.Scope = Scope;
                                        PF.dataCache = dataCache;

                                        tmpProfit = PF.ProfitFunctionT();

                                        SignalsNum = PF.SignalsNum;
                                        if (tmpProfit > BestChrom.J || Generation == 0)
                                        {
                                            BestChrom.J = tmpProfit;
                                            BestChrom.Genes = Gens;
                                            BestChrom.SignalsNum = SignalsNum;
                                        }
                                        List<Chromosome> tmpList = new List<Chromosome>();
                                        tmpList.Add(BestChrom);
                                        XLS.ExcelReportT_Overal(tmpList, Generation + 1, 1, xlPackage, Generation, Symb.Symbol);
                                    }
                                    catch (Exception e)
                                    {
                                        XLS.ExcelReport_Log(e.Message, xlPackage, Generation, Symb.Symbol);
                                    }
                                    Generation++;

                                }
                            }
                        }
                    }
                    break;
            }



            return BestChrom;
        }

        //Greedy Grid Algorithm
        public Chromosome GreedyGrid(int GridNum)
        {
            Chromosome BestChrom = new Chromosome();
            List<Chromosome> ChrmList;
            List<double> Gens = new List<double>();
            int Step = GridNum;


            BestChrom.Genes.Add(0);
            BestChrom.Genes.Add(0);
            BestChrom.Genes.Add(0);
            BestChrom.Genes.Add(0);
            BestChrom.Genes.Add(0);
            BestChrom.Genes.Add(0);
            BestChrom.Genes.Add(0);
            BestChrom.Genes.Add(0);

            Generation = 0;

            
            int notchanged = 0;

            List<List<double>> PRange = new List<List<double>>();

            List<double> RRange;

            foreach (var a in ChromRange)
            {
                RRange = new List<double>();
                RRange.Add(a[0]);
                RRange.Add(a[1]);

                PRange.Add(RRange);

            }
            int SignalsNum = 0;
            while (notchanged == 0)
            {
                ChrmList = new List<Chromosome>();

                switch (Indicator)
                {
                    case "Alligator":
                        foreach (var p0 in PRange[0])
                        {
                            foreach (var p1 in PRange[1])
                            {
                                foreach (var p2 in PRange[2])
                                {
                                    foreach (var p3 in PRange[3])
                                    {
                                        foreach (var p4 in PRange[4])
                                        {
                                            foreach (var p5 in PRange[5])
                                            {
                                                foreach (var p6 in PRange[6])
                                                {
                                                    foreach (var p7 in PRange[7])
                                                    {
                                                        List<double> Genes = new List<double>();
                                                        Genes.Add((int)p0);
                                                        Genes.Add((int)p1);
                                                        Genes.Add((int)p2);
                                                        Genes.Add((int)p3);
                                                        Genes.Add((int)p4);
                                                        Genes.Add((int)p5);
                                                        Genes.Add((int)p6);
                                                        Genes.Add(p7);
                                                        
                                                        double TProfit;
                                                        PF = new ProfitFunction();
                                                        PF.Indicator = Indicator;
                                                        PF.InsCode = Symb.InsCode;
                                                        PF.SDEven = SDEven;
                                                        PF.SHEven = SHEven;
                                                        PF.EDEven = EDEven;
                                                        PF.EHEven = EHEven;
                                                        PF.Genes = Gens;
                                                        PF.Scope = Scope;
                                                        PF.dataCache = dataCache;

                                                        TProfit = PF.ProfitFunctionT();

                                                        SignalsNum = PF.SignalsNum;

                                                        Chromosome tmpChrm = new Chromosome();
                                                        tmpChrm.Genes = Genes;
                                                        tmpChrm.J = TProfit;
                                                        tmpChrm.SignalsNum = SignalsNum;
                                                        ChrmList.Add(tmpChrm);
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        break;
                }
                ChrmList = ChrmList.OrderByDescending(t => t.J).ToList();
                BestChrom = ChrmList[0];

                List<Chromosome> tmpchr = new List<Chromosome>();
                tmpchr.Add(BestChrom);

                XLS.ExcelReportT_Overal(tmpchr, Generation + 1, 1, xlPackage, Generation, Symb.Symbol);

                notchanged = 1;
                int i = 0;
                foreach (var a in ChrmList[0].Genes)
                {

                    if (PRange[i][0] != PRange[i][1] && ChromRange[i][2] == 1) notchanged = 0;

                    if (a == PRange[i][0])
                    {
                        if (ChromRange[i][2] == 1)
                        {
                            PRange[i][1] = Math.Floor((PRange[i][1] + PRange[i][0]) / 2);
                        }
                        else
                        {
                            PRange[i][1] = (PRange[i][1] + PRange[i][0]) / 2;
                        }

                    }
                    else
                    {
                        if (ChromRange[i][2] == 1)
                        {
                            PRange[i][0] = Math.Ceiling((PRange[i][1] + PRange[i][0]) / 2);
                        }
                        else
                        {
                            PRange[i][0] = (PRange[i][1] + PRange[i][0]) / 2;
                        }
                    }
                    i++;
                }
                Generation++;
            }


            return BestChrom;
        }
    }
}
