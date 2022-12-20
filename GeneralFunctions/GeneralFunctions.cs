using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TabularEditor.TOMWrapper;
using TabularEditor.Scripting;


namespace GeneralFunctions
{
    public static class Fx
    {
        static readonly Model Model;
        static readonly TabularEditor.UI.UITreeSelection Selected;

        public static void Error(string message, int lineNumber = -1, bool suppressHeader = false)
        {
            ScriptHelper.Error(message: message, lineNumber: lineNumber, suppressHeader: suppressHeader);
        }

        public static void Info(string message, int lineNumber = -1)
        {
            ScriptHelper.Info(message: message, lineNumber: lineNumber);
        }

        public static Table SelectTable(IEnumerable<Table> tables, Table preselect, string label)
        {
            return ScriptHelper.SelectTable(tables: tables, preselect: preselect, label: label);
        }

        public static void Output(object value, int lineNumber = -1)
        {
            ScriptHelper.Output(value: value, lineNumber: lineNumber);
        }

        //public static class Fx { 

        public static Table CreateCalcTable(
        string tableName = "Measures",
        string tableExpression = "FILTER({BLANK()},FALSE)")
        {
            if (Model.Tables.Any(t => t.Name == tableName))
            {
                Error("There is already a table called " + tableName + ".");
                return null;
            };

            Table newTable =
                Model.AddCalculatedTable(
                    name: tableName,
                    expression: tableExpression);

            return newTable;

        }



        //Func<Table, bool> mTables = (t) => !model.Relationships.Any(r => r.ToTable == t || r.FromTable == t)
        //    //the table has no visible columns
        //    && t.Columns.All(c => !c.IsVisible);


        public static Table GetMeasureTable(string tableName = "Measures", bool createIfNecessary = true)
        {


            var MeasureTables = Model.Tables.Where
                    (t =>
                        //no relationships start or end on the table
                        !Model.Relationships.Any(r => r.ToTable == t || r.FromTable == t)
                        //the table has no visible columns
                        && t.Columns.All(c => !c.IsVisible)
                    );
            if (MeasureTables.Count() == 0)
            {
                if (createIfNecessary)
                {
                    //does not exist, new one is created
                    Table newMeasureTable =
                        Model.AddCalculatedTable(
                            name: tableName,
                            expression: "FILTER({BLANK()},FALSE)"
                        );
                    return newMeasureTable;

                }
                else
                {
                    //does not exist, will not create
                    return null;
                };
            }
            else if (MeasureTables.Count() == 1)
            {
                Table singleExistingTable = MeasureTables.First();
                return singleExistingTable;
            }
            else
            {
                Table chosenTable = ScriptHelper.SelectTable(MeasureTables, MeasureTables.First(), "Select measure table to use");
            }

            return null;
        }


    }
}
