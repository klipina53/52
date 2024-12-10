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
    public class GroupContext : Group
    {
        public GroupContext(int Id, string Name) : base(Id, Name) { }
        public static List<GroupContext> AllGroups()
        {
            List<GroupContext> allGroups = new List<GroupContext>();
            MySqlConnection connection = Connection.OpenConnection();
            MySqlDataReader BDGroups = Connection.Query("SELECT * FROM `Group` ORDER BY `Name`;", connection);
            while (BDGroups.Read())
            {
                allGroups.Add(new GroupContext(
                    BDGroups.GetInt32(0),
                    BDGroups.GetString(1)));
            }
            Connection.CloseConnection(connection);
            return allGroups;
        }
    }
}
