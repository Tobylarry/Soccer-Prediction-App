using System;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace BettingApplication2
{

    public class Program
    {
        public static void getExcelFile()
            {
            decimal starting_bankroll = 100;
            decimal bankroll = starting_bankroll;
            decimal upsets = 0;
            decimal non_upsets = 0;
            int underdog_upsets = 0;
            int away_non_upsets = 0;
            int initial_amount = 100;
            int wager_size = 5;
            int amount = wager_size;
            string home_teams = "";
            string away_teams = "";
            int home_goals = 0;
            int away_goals = 0;
            decimal home_odds = 0;
            decimal draw_odds = 0;
            decimal away_odds = 0;
            double ROI = 0;
            int draws = 0;
            double PercentageOfUpsets = 0;
            int TotalMatches = 0;
            int homewins = 0;
            int awaywins = 0;
           
            string path = "C:/Users/Tobyl/Desktopp/project_comp/tobi.csv";

                    string[] lines = System.IO.File.ReadAllLines(path);
                    foreach(string line in lines)
                    {
                        var rowdetails = line.Split(',');
                            if(rowdetails[0] == "Div")
                            {
                               continue;
                            }
                            else
                            {
                              home_teams = rowdetails[2];
                              away_teams = rowdetails[3];
                              home_goals = Convert.ToInt32(rowdetails[4]);
                              away_goals = Convert.ToInt32(rowdetails[5]);
                              home_odds  = Convert.ToDecimal(rowdetails[23]);
                              draw_odds  = Convert.ToDecimal(rowdetails[24]);
                              away_odds  = Convert.ToDecimal(rowdetails[25]);
                                     if (home_odds > away_odds)
                                     {
                                           if(home_goals > away_goals)
                                           {
                                                  upsets += 1;
                                                  bankroll += (wager_size * (home_odds - 1));
                                           }
                                           else
                                           {
                                                non_upsets += 1;
                                                bankroll -= wager_size;
                                           }

                                     }
                                     if(home_goals == away_goals)
                                     {
                                      draws += 1;

                                     }
                                     if (home_goals > away_goals)
                                     {
                                       homewins += 1;
                                     }
                                    else if (home_goals < away_goals)
                                    {
                                        awaywins += 1;
                                    }
                }
                      
                    }
                       TotalMatches = (draws + homewins + awaywins);
                       PercentageOfUpsets = Convert.ToDouble(upsets/TotalMatches) * 100;
                       PercentageOfUpsets = Math.Round(PercentageOfUpsets, 2);
                       ROI = Convert.ToDouble((bankroll - starting_bankroll) / (wager_size * TotalMatches)) * 100;
                       Console.WriteLine($"There were {upsets} upsets in a total of {TotalMatches} football matches");

                       Console.WriteLine($"There were {draws} draws in a total of {TotalMatches} football matches");

                       Console.WriteLine($"There were {homewins} homewins in a total of {TotalMatches} football matches");

                       Console.WriteLine($"There were {awaywins} awaywins in a total of {TotalMatches} football matches");

                       Console.WriteLine($"Therefore the perecentage of upset is {PercentageOfUpsets} %");

                       Console.WriteLine($"Starting bankroll : {starting_bankroll}");

                       Console.WriteLine($"Finsihing bankroll : {bankroll}");

                       Console.WriteLine($"ROI = {ROI}");

                       Console.ReadLine();
        }

        public static void Main(string[] args)
        {
            getExcelFile();
            
        }
    }

           
}





    
    


