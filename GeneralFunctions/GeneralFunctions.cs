using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TabularEditor.TOMWrapper;
using TabularEditor.Scripting;
using System.Reflection.Emit;

namespace GeneralFunctions
{

    //copy from the following line up to ****** and remove the // before the closing bracket
    public static class Fx
    {
        
        //in TE2 (at least up to 2.17.2) any method that accesses or modifies the model needs a reference to the model 
        //the following is an example method where you can build extra logic
        public static Table CreateCalcTable(Model model, string tableName, string tableExpression) 
        { 
            return model.AddCalculatedTable(name:tableName,expression:tableExpression);
        }

        public static Table SelectTableExt(Model model, string possibleName = null, string annotationName = null, string annotationValue = null, 
            Func<Table,bool>  lambdaExpression = null, string label = "Select Table", bool skipDialogIfSingleMatch = true, bool showOnlyMatchingTables = true)
        {
            
            if (lambdaExpression == null)
            {
                if (possibleName != null) { 
                    lambdaExpression = (t) => t.Name == possibleName;
                } else if(annotationName!= null && annotationValue != null)
                {
                    lambdaExpression = (t) => t.GetAnnotation(annotationName) == annotationValue;
                }
            }

            IEnumerable<Table> tables = model.Tables.Where(lambdaExpression);

            //none found, let the user choose from all tables
            if (tables.Count() == 0)
            {
                return SelectTable(tables: model.Tables, label: label);
                
            }
            else if (tables.Count() == 1 && !skipDialogIfSingleMatch)
            {
                return SelectTable(tables: model.Tables, preselect: tables.First(), label: label);
            }
            else if (tables.Count() == 1 && skipDialogIfSingleMatch)
            {
                return tables.First();
            } 
            else if (tables.Count() > 1 && showOnlyMatchingTables)
            {
                return SelectTable(tables: tables, preselect: tables.First(), label: label);
            }
            else if (tables.Count() > 1 && !showOnlyMatchingTables)
            {
                return SelectTable(tables: model.Tables, preselect: tables.First(), label: label);
            } else
            {
                Error(@"Unexpected logic in ""SelectTableExt""");
                return null;
            }
        }
        
        //add other methods always as "public static" followed by the data type they will return or void if they do not return anything.



        //}

        //******************
        //do not copy from this line below, and remove the // before the closing bracket above to close the class definition


        //Model and Selected cannot be accessed directly. Always pass a reference to the requited objects. 
        //static readonly Model Model;
        //static readonly TabularEditor.UI.UITreeSelection Selected;


        //These functions replicate the ScriptHelper functions so that they can be
        //used inside the script without the ScriptHelper prefix which cannot be used inside tabular editor
        //the list is not complete and does not include all overloads, complete as necessary. 
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

        public static void Output(object value, int lineNumber = -1)
        {
            ScriptHelper.Output(value: value, lineNumber: lineNumber);
        }
    }
}
