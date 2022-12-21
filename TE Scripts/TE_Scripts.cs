using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TabularEditor.TOMWrapper;
using TabularEditor.Scripting;
//using GeneralFunctions;

namespace TE_Scripting
{
    public class TE_Scripts
    {

        void createLabelCalcGroup()
        {
            String calcGroupName = "Labels 2";

            CalculationGroupTable cg = Model.AddCalculationGroup(name: calcGroupName);

            Table t = SelectTable(Model.Tables, label: "Select table of axis field");
            if (t == null) return;

            Column c = SelectColumn(table: t, label: "Select axis field");
            if (c == null) return;

            String calcItemName = "Only Top Value by " + c.Name;

            String calcItemExpression =
                String.Format(
                    @"VAR maxValue =
                        CALCULATE(
                            MAXX(
                                VALUES(
                                    {0}
                                ),
                                SELECTEDMEASURE( )
                            ),
                            ALLSELECTED( {1} )
                        )
                    VAR currentValue =
                        SELECTEDMEASURE( )
                    VAR fString =
                        IF(
                            maxValue = currentValue,
                            SELECTEDMEASUREFORMATSTRING(

                            ),
                            "";;;""
                        )
                    RETURN
                        fString", c.DaxObjectFullName, t.DaxObjectFullName);

            CalculationItem ci = cg.AddCalculationItem(name: calcItemName, expression: calcItemExpression);
            ci.FormatDax();


        }


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

        public static Table SelectTable(IEnumerable<Table> tables, Table preselect = null, string label = "Select Table")
        {
            return ScriptHelper.SelectTable(tables: tables, preselect: preselect, label: label);
        }

        public static Column SelectColumn(Table table, Column preselect = null, string label = "Select Column")
        {
            return ScriptHelper.SelectColumn(table: table, preselect: preselect, label: label); 
        }

        public static void Output(object value, int lineNumber = -1)
        {
            ScriptHelper.Output(value: value, lineNumber: lineNumber);
        }


    }
}
