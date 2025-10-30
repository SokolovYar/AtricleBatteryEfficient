using ClosedXML.Excel;
using Google.OrTools.LinearSolver;
using MCalc;
using System;


var battery = new BatteryCore(
            emax: 100,     
            emin: 0,       
            pCharge: 30,   
            pDischarge: 30,
            etaCharge: 0.95,
            etaDischarge: 0.95
        );


DataReader dataReader = new DataReader();


var model = new BatteryEfficiency(battery, dataReader.Read());
model.Optimize();

class BatteryCore
{
    public double Ecurrent { get; set; } = 0;
    public double Emax { get; set; }
    public double Emin { get; set; }
    public double PCharge { get; set; }
    public double PDischarge { get; set; }
    public double EtaCharge { get; set; }
    public double EtaDischarge { get; set; }

    public BatteryCore(double emax, double emin, double pCharge, double pDischarge, double etaCharge, double etaDischarge)
    {
        Emax = emax;
        Emin = emin;
        PCharge = pCharge;
        PDischarge = pDischarge;
        EtaCharge = etaCharge;
        EtaDischarge = etaDischarge;
    }
}

class BatteryEfficiency
{
    private BatteryCore _battery;
    private Solver solver;

    public List<double> Prices { get; set; } 

    public BatteryEfficiency(BatteryCore battery, List<double> prices)
    {
        _battery = battery;
        Prices = prices;
        solver = Solver.CreateSolver("GLOP"); // Linear Google solver
    }

    public void Optimize()
    {
        int T = Prices.Count;

        var Pcharge = new Variable[T];
        var Pdisch = new Variable[T];
        var E = new Variable[T + 1]; 

        for (int t = 0; t < T; t++)
        {
            Pcharge[t] = solver.MakeNumVar(0, _battery.PCharge, $"Pch_{t}");
            Pdisch[t] = solver.MakeNumVar(0, _battery.PDischarge, $"Pdis_{t}");
        }

        for (int t = 0; t <= T; t++)
        {
            E[t] = solver.MakeNumVar(_battery.Emin, _battery.Emax, $"E_{t}");
        }

        // Initial condition of charge
        solver.Add(E[0] == _battery.Ecurrent);

        // Energy Balance
        for (int t = 0; t < T; t++)
        {
            solver.Add(
                E[t + 1] == E[t]
                + _battery.EtaCharge * Pcharge[t]
                - (1.0 / _battery.EtaDischarge) * Pdisch[t]
            );
        }

        // === Objective function (profit maximization) ===
        Objective objective = solver.Objective();
        for (int t = 0; t < T; t++)
        {
            objective.SetCoefficient(Pdisch[t], Prices[t]); 
            objective.SetCoefficient(Pcharge[t], -Prices[t]); 
        }
        objective.SetMaximization();

        // === Solving ===
        Solver.ResultStatus resultStatus = solver.Solve();

        if (resultStatus == Solver.ResultStatus.OPTIMAL)
        {
            Console.WriteLine($"Maximum profit = {solver.Objective().Value():F2}");
            Console.WriteLine("t\tPrice\tCharge\tDisch\tEnergy");

            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Results");

            ws.Cell(1, 1).Value = "t";
            ws.Cell(1, 2).Value = "Price";
            ws.Cell(1, 3).Value = "Pcharge";
            ws.Cell(1, 4).Value = "Pdisch";
            ws.Cell(1, 5).Value = "Energy";

            for (int t = 0; t < T; t++)
            {
                double eValue = E[t].SolutionValue();

                Console.WriteLine($"{t}\t{Prices[t],5:F1}\t{Pcharge[t].SolutionValue(),6:F2}\t{Pdisch[t].SolutionValue(),6:F2}\t{eValue,6:F2}");

                ws.Cell(t + 2, 1).Value = t;
                ws.Cell(t + 2, 2).Value = Prices[t];
                ws.Cell(t + 2, 3).Value = Pcharge[t].SolutionValue();
                ws.Cell(t + 2, 4).Value = Pdisch[t].SolutionValue();
                ws.Cell(t + 2, 5).Value = eValue;
            }

            // === Calculation of equivalent cycles ===
            double charged = 0, discharged = 0;
            double tol = 1e-6;

            for (int t = 0; t < T; t++)
            {
                double diff = E[t + 1].SolutionValue() - E[t].SolutionValue();
                if (diff > tol) charged += diff;
                else if (diff < -tol) discharged += -diff;
            }

            // Convert to equivalent cycles
            double chargedCycles = charged / _battery.Emax;
            double dischargedCycles = discharged / _battery.Emax;
            double totalCycles = (charged + discharged) / (2.0 * _battery.Emax);

            Console.WriteLine($"\nCharged energy = {charged:F3}");
            Console.WriteLine($"Discharged energy = {discharged:F3}");
            Console.WriteLine($"Equivalent cycles (round-trip) = {totalCycles:F4}");

            int summaryRow = T + 3;
            ws.Cell(summaryRow, 1).Value = "Charged energy";
            ws.Cell(summaryRow, 2).Value = charged;
            ws.Cell(summaryRow + 1, 1).Value = "Discharged energy";
            ws.Cell(summaryRow + 1, 2).Value = discharged;
            ws.Cell(summaryRow + 2, 1).Value = "Equivalent cycles (round-trip)";
            ws.Cell(summaryRow + 2, 2).Value = totalCycles;

            ws.Columns().AdjustToContents();
            wb.SaveAs("Results.xlsx");
            Console.WriteLine("Results saved to Results.xlsx");
        }
        else
        {
            Console.WriteLine("No solution found.");
        }
    }
}




