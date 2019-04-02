# 'Control-M/Agent Base' Class

## Introduction

This class is designed to contain some basic types of the Control-M Agent Automatic Reception Application.

## Version 0.0.0.1

![Debug 1](C:\Users\MoChen\source\repos\ControlMAgent.Base\ControlMAgent.Base\readme\Debug 1.bmp)

This dll can do this. Current Development note.

Okay, it's a success, now we can correctly get the next single row.

![1554185051881](C:\Users\MoChen\source\repos\ControlMAgent.Base\ControlMAgent.Base\readme\1554185051881.png)

Next is to extend next area. The extension should be 4 columns.

![1554187215662](C:\Users\MoChen\source\repos\ControlMAgent.Base\ControlMAgent.Base\readme\1554187215662.png)

Successful. Now Plug this into the main application.

Below is the base code for version 0.0.0.1:

```C#
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
        string NextArea(int RowsDown);
    }
    //Summary
    //  A CellRange Object is an object that returns an "expression" when called upon.
    //  
    public class CellRange : ICellRange
    {
        private readonly string initRange;
        private string currentRange;
        private string startRange;
        private string endinRange;
        private bool isArea;

        private char ColS, ColE;
        private int RowS, RowE;
        public CellRange(string InitialCell)
        {
            //is it a valid string? "$C$15", or "$C$15:$D$16"
            Regex rxCellString_Single = new Regex(@"^\$\w+\$\d+$");
            Regex rxCellString_Area = new Regex(@"^\$\w+\$\d+:\$\w+\$\d+$");
            isArea = true;
            if (!rxCellString_Single.IsMatch(InitialCell) && !rxCellString_Area.IsMatch(InitialCell))
                throw new FormatException(InitialCell + " is not a valid Cell Expression");

            initRange = InitialCell.Contains(':') ? InitialCell : string.Format("{0}:{0}", InitialCell);
            //"$C$15:$D$16"
            Regex rxFetch = new Regex(@"^\$(?'C0'\w)\$(?'R0'\d+):\$(?'C1'\w)\$(?'R1'\d+)$");
            Match rxMatch = rxFetch.Match(initRange);
            ColS = Convert.ToChar(rxMatch.Groups["C0"].Value);
            RowS = Convert.ToInt32(rxMatch.Groups["R0"].Value);
            ColE = Convert.ToChar(rxMatch.Groups["C1"].Value);
            RowE = Convert.ToInt32(rxMatch.Groups["R1"].Value);
            // Add code, if ColS = ColE andd RowS = RowE then is Area = false
            if (ColS == ColE && RowS == RowE) isArea = false;
            //Below in Case of "$C$15:$D$16"
            currentRange = string.Format(@"{0}:{1}"
                , Comb_Cell(ColS, RowS), Comb_Cell(ColE, RowE)); //$C$15:$C$15
            startRange = Comb_Cell(ColS, RowS); //$C$15
            endinRange = Comb_Cell(ColS, RowE); //$C$16
        }
        //Combine a Char and a Int into a $C$5
        private string Comb_Cell(char col, int row) => string.Format("${0}${1}", col, row);
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
                        ExtendedResult += Comb_Cell(c, i) + ";";
                    }
                }
                return ExtendedResult;
            }
            else
            {
                return endinRange;
            }
        }

        //Itinerate one row down to the next row
        public string NextRow()
        {
            //throw new NotImplementedException();
            //When the object is area, return the next single row in the last row
            Match rxMatch = new Regex(@"^\$(?'C0'\w+)\$(?'R0'\d+)$").Match(endinRange);
            char Col = Convert.ToChar(rxMatch.Groups["C0"].Value);
            int Row = Convert.ToInt32(rxMatch.Groups["R0"].Value);
            Row += 1;
            endinRange = startRange = Comb_Cell(Col, Row);
            currentRange = string.Format(@"{0}:{0}", endinRange);
            return endinRange;
        }

        public string NextArea(int RowsDown)
        {
            //throw new NotImplementedException();
            if (RowsDown < 1)
                throw new InvalidOperationException("Argument RowsDown must be greater than 1.");
            Match rxMatch = new Regex(@"^\$(?'C0'\w+)\$(?'R0'\d+)$").Match(endinRange);
            char Col = Convert.ToChar(rxMatch.Groups["C0"].Value);
            int Row = Convert.ToInt32(rxMatch.Groups["R0"].Value);
            char ColE = Convert.ToChar(Convert.ToInt32(Col) + 3);
            int RowE = Row + RowsDown;
            int RowS = Row + 1;
            isArea = true;
            currentRange = string.Format(@"{0}:{1}"
                , Comb_Cell(Col, RowS), Comb_Cell(ColE, RowE));
            startRange = Comb_Cell(Col, RowS);
            endinRange = Comb_Cell(Col, RowE);
            return Extend();
        }
    }
}

```

