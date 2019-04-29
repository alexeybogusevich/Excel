using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace ExcelApplication
{
    public class MyCell : DataGridViewTextBoxCell
    {
        double val; //результат виразуS
        string name; //ім'я
        string exp; //вираз
        private int row;
        private int column;
        private List<string> cellsIDependOn = new List<string>();
        private List<string> cellsDependentOnMe = new List<string>();


        public MyCell()
        {
            name = "A0";
            val = 0;
            exp = "0";
        }

        public MyCell(string n)
        {
            name = n;
            val = 0;
            exp = "0";
        }

        public string Name
        {
            get { return name; }
        }

        public double Value
        {
            get { return val; }
            set { val = value; }
        }

        public string Exp
        {
            get { return exp; }
            set { exp = value; }
        }

        public int Row
        {
            get { return row; }
            set { row = value; }
        }

        public int Column
        {
            get { return column; }
            set { column = value; }
        }

        public List<string> IDependOn
        {
            get { return cellsIDependOn; }
            set { cellsIDependOn = value; }
        }

        public List<string> DependentOnMe
        {
            get { return cellsDependentOnMe; }
            set { cellsDependentOnMe = value; }
        }

        public string AddressToValue(ref Dictionary<string, MyCell> dictionary)
        {
            string expression = Exp;
            bool address = false;
            bool operators = false;
            int minmax = 0;
            int addressStart = -1;
            int addressEnd = -1;
            string allTheOperators = "+-/*%";

            if (!AddressPresence(expression))
                return expression;

            try
            {
                for (int i = 0; i < expression.Length; ++i)
                {
                    if (minmax>0)
                    {
                        if (expression[i] == ';')
                            continue;
                        else if (expression[i] == ')')
                        {
                            --minmax;
                            continue;
                        }
                    }
                    if (Char.IsSeparator(expression[i]))
                        continue;
                    else if (allTheOperators.Contains(expression[i]))
                    {
                        if (!operators)
                            operators = true;
                        else
                            throw new Exception();
                    }
                    else if (Char.IsUpper(expression[i]))
                    {
                        if (operators)
                            operators = false;
                        if ((i+3)<=expression.Length && (expression.Substring(i, 3) == "Min" || expression.Substring(i, 3) == "Max"))
                        {
                            i += 3;
                            ++minmax;
                            continue;
                        }
                        if (((i + 1) <= expression.Length) && (!Char.IsLetterOrDigit(expression[i + 1])
                            || i + 1 == expression.Length))
                            throw new Exception();
                        if (!address)
                        {
                            address = true;
                            addressStart = i;
                            addressEnd = i;
                        }
                        else
                        {
                            ++addressEnd;
                        }
                    }
                    else if (Char.IsDigit(expression[i]) && address)
                    {
                        if (operators)
                            operators = false;
                        ++addressEnd;
                        bool ok = false;
                        if (i + 1 != expression.Length)
                        {
                            if (Char.IsDigit(expression[i + 1]))
                                continue;
                            else if (Char.IsSeparator(expression[i + 1]))
                                ok = true;
                            else if (allTheOperators.Contains(expression[i + 1]))
                                ok = true;
                            else if (minmax>0 && (expression[i + 1] == ';' || expression[i + 1] == ')'))
                                ok = true;
                            else
                                throw new Exception();
                        }
                        else
                            ok = true;
                        if (ok)
                        {
                            string cellName = expression.Substring(addressStart, addressEnd - addressStart + 1);
                            if(!IDependOn.Contains(cellName))
                                IDependOn.Add(cellName);
                            if(!dictionary[cellName].DependentOnMe.Contains(this.Name))
                                dictionary[cellName].DependentOnMe.Add(this.Name);
                            expression = expression.Replace(cellName, dictionary[cellName].Value.ToString());
                            address = false;
                        }
                    }
                    else if(Char.IsDigit(expression[i]) && !address)
                    {
                        if (operators)
                            operators = false;
                    }
                }
            }
            catch(Exception)
            {
                MessageBox.Show("Invalid Expression\nThe cell will be cleared.");
                Exp = "0";
                return null;
            }
            return expression;
        } //перетворяє рядок на числовий вираз

        private bool AddressPresence(string text)
        {
            for (int i = 0; i < text.Length; ++i)
            {
                if (Char.IsUpper(text[i]))
                {
                    if ((i + 3) <= text.Length && (text.Substring(i, 3) == "Min"
                        || text.Substring(i, 3) == "Max"))
                    {
                        i += 3;
                        continue;
                    }
                    return true;
                }
            }
            return false;
        } //перевіряє на наявність адрес в рядку

    }
}
