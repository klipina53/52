﻿using MySql.Data.MySqlClient;
using ReportGeneration_Lipina.Common;
using ReportGeneration_Lipina.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportGeneration_Lipina.Classes
{
    public class EvaluationContext : Evaluation
    {
        public EvaluationContext(int Id, int IdWork, int IdStudent, string Value, string Lateness) :
            base(Id, IdWork, IdStudent, Value, Lateness) { }
        public static List<EvaluationContext> AllEvaluations()
        {
            List<EvaluationContext> allEvaluations = new List<EvaluationContext>();
            MySqlConnection connection = Connection.OpenConnection();
            MySqlDataReader BDEvaluations = Connection.Query("SELECT * FROM `Evaluation`;", connection);
            while (BDEvaluations.Read())
            {
                allEvaluations.Add(new EvaluationContext(
                    BDEvaluations.GetInt32(0),
                    BDEvaluations.GetInt32(1),
                    BDEvaluations.GetInt32(2),
                    BDEvaluations.GetString(3),
                    BDEvaluations.GetString(4)));
            }
            Connection.CloseConnection(connection);
            return allEvaluations;
        }
    }
}
