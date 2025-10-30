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
dataReader.Read();

//var model = new BatteryEfficiency(battery);
//model.Optimize();

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

    public List<double> Prices { get; set; } = new List<double> {
        20, 18, 15, 14, 13, 12, 14, 18, 25, 30, 35, 40,
        38, 36, 34, 32, 30, 28, 26, 24, 22, 21, 20, 19
    };

    public BatteryEfficiency(BatteryCore battery)
    {
        _battery = battery;
        solver = Solver.CreateSolver("GLOP"); // линейный солвер
    }

    public void Optimize()
    {
        int T = Prices.Count;

        // === Переменные ===
        var Pcharge = new Variable[T];
        var Pdisch = new Variable[T];
        var E = new Variable[T + 1]; // E[0]...E[T]

        for (int t = 0; t < T; t++)
        {
            Pcharge[t] = solver.MakeNumVar(0, _battery.PCharge, $"Pch_{t}");
            Pdisch[t] = solver.MakeNumVar(0, _battery.PDischarge, $"Pdis_{t}");
        }

        for (int t = 0; t <= T; t++)
        {
            E[t] = solver.MakeNumVar(_battery.Emin, _battery.Emax, $"E_{t}");
        }

        // === Начальное условие ===
        solver.Add(E[0] == _battery.Ecurrent);

        // === Баланс энергии ===
        for (int t = 0; t < T; t++)
        {
            solver.Add(
                E[t + 1] == E[t]
                + _battery.EtaCharge * Pcharge[t]
                - (1.0 / _battery.EtaDischarge) * Pdisch[t]
            );
        }

        // === Целевая функция (максимизация прибыли) ===
        Objective objective = solver.Objective();
        for (int t = 0; t < T; t++)
        {
            objective.SetCoefficient(Pdisch[t], Prices[t]); // прибыль от продажи
            objective.SetCoefficient(Pcharge[t], -Prices[t]); // затраты на зарядку
        }
        objective.SetMaximization();

        // === Решение ===
        Solver.ResultStatus resultStatus = solver.Solve();

        if (resultStatus == Solver.ResultStatus.OPTIMAL)
        {
            Console.WriteLine($"Максимальная прибыль = {solver.Objective().Value():F2}");
            Console.WriteLine("t\tPrice\tCharge\tDisch\tEnergy");

            for (int t = 0; t < T; t++)
            {
                Console.WriteLine($"{t}\t{Prices[t],5:F1}\t{Pcharge[t].SolutionValue(),6:F2}\t{Pdisch[t].SolutionValue(),6:F2}\t{E[t].SolutionValue(),6:F2}");
            }
        }
        else
        {
            Console.WriteLine("Решение не найдено.");
        }
    }
}




