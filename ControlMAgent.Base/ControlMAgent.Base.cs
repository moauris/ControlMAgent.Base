using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ControlMAgent.Base
{
    internal interface ICellRange
    {
        //The Current Cell Initial Expression, such as "$C$15", or "$C$15:$D$16"
        string InitRange();
        //used to signal Iteration with Next()
        string CurrentRange();
        //The Current Cell Full String, "$C$15" => "$C$15"
        //or in case of area, "$C$15:$D$16" => "$C$15;$C$16;$D$15;$D$17"
        string Extend();
        //Return if a cell is an area or a single cell
        bool IsArea();
        //Return the Current Starting Cell
        string StartRange();
        //Return the Current Ending Cell
        string EndinRange();
        //Return the next row
        string NextRow();
        //Return the next area
        string NextArea();
    }
    //Summary
    //  A CellRange Object is an object that returns an "expression" when called upon.
    //  
    public class CellRange : ICellRange
    {
        private readonly string initRange;
        private readonly string currentRange;
        private readonly string startRange;
        private readonly string endinRange;
        private readonly bool isArea;
        public CellRange(string InitialCell)
        {
            //is it a valid string? "$C$15", or "$C$15:$D$16"
            Regex rxCellString_Single = new Regex(@"^\$\w+\$\d+$");
            Regex rxCellString_Area = new Regex(@"^\$\w+\$\d+:\$\w+\$\d+$");
            isArea = false;
            if (!rxCellString_Single.IsMatch(InitialCell))
            {
                if (!rxCellString_Area.IsMatch(InitialCell))
                {
                    throw new FormatException(InitialCell + " is not a valid Cell Expression");
                }
                isArea = true;
            }
            initRange = InitialCell;
            currentRange = InitialCell;
            startRange = InitialCell;
            endinRange = InitialCell;
        }
        public string InitRange() => initRange;
        public string CurrentRange() => currentRange;
        public string StartRange() => startRange;
        public string EndinRange() => endinRange;
        public bool IsArea() => isArea;
        public string Extend()
        {
            if (isArea)
            {
                //"$C$15:$D$16"
                Regex rxFetch = 
                    new Regex(@"^\$(?'C0'\w)\$(?'R0'\d+):\$(?'C1'\w)\$(?'R1'\d+)$");
                Match rxMatch = rxFetch.Match(currentRange);
                char ColStart = Convert.ToChar(rxMatch.Groups["C0"].Value);
                int RowStart = Convert.ToInt32(rxMatch.Groups["R0"].Value);
                char ColEnd = Convert.ToChar(rxMatch.Groups["C1"].Value);
                int RowEnd = Convert.ToInt32(rxMatch.Groups["R1"].Value);
                string ExtendedResult = "";
                for(char c = ColStart; c <= ColEnd; c++)
                {
                    for (int i = RowStart; i <= RowEnd; i++)
                    {
                        ExtendedResult += string.Format(@"${0}${1};", c, i);
                    }
                }
                return ExtendedResult;
            }
            else
            {
                return currentRange;
            }
        }


        public string NextRow()
        {
            throw new NotImplementedException();
        }

        public string NextArea()
        {
            throw new NotImplementedException();
        }




    }
}
