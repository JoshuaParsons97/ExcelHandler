using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace Interface
{
    internal class TestData
    {
        public static DataSet TestDataSet
        {
            get
            {
                DataSet DataSet = new DataSet();
                ConstructDataSet(ref DataSet);
                return DataSet;
            }  
        }

        private static void ConstructDataSet(ref DataSet DataSet)
        {
            DataSet TestData = new();

            DataTable Table1 = new("Users");
            Table1.Columns.Add("Name");
            Table1.Columns.Add("Surname");
            Table1.Columns.Add("Email");
            Table1.Columns.Add("Address");
            TestData.Tables.Add(Table1);

            DataTable Table2 = new("Accounts");
            Table2.Columns.Add("Account Number");
            Table2.Columns.Add("Account Holder");
            Table2.Columns.Add("Date Opened");
            Table2.Columns.Add("Status");
            TestData.Tables.Add(Table2);
        }
    }
}
