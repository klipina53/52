using MySql.Data.MySqlClient;
using ReportGeneration_Lipina.Common;
using ReportGeneration_Lipina.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportGeneration_Lipina.Classes
{
    public class WorkContext : Work
    {
        public WorkContext(int Id, int IdDiscipline, int IdType, DateTime Date, string Name, int Semester) :
            base(Id, IdDiscipline, IdType, Date, Name, Semester) { }
        public static List<WorkContext> AllWorks()
        {
            List<WorkContext> allWorks = new List<WorkContext>();
            MySqlConnection connection = Connection.OpenConnection();
            MySqlDataReader BDWorks = Connection.Query("SELECT * FROM `Work` ORDER BY `Date`;", connection);
            while (BDWorks.Read())
            {
                allWorks.Add(new WorkContext(
                    BDWorks.GetInt32(0),
                    BDWorks.GetInt32(1),
                    BDWorks.GetInt32(2),
                    BDWorks.GetDateTime(3),
                    BDWorks.GetString(4),
                    BDWorks.GetInt32(5)));
            }
            Connection.CloseConnection(connection);
            return allWorks;
        }
    }
}
