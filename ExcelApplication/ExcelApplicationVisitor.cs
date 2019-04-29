using System;
using System.Collections.Generic;
using System.Diagnostics;
using Antlr4;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Antlr4.Runtime.Misc;

namespace ExcelApplication
{
    class ExcelApplicationVisitor : CombinedBaseVisitor<double>
    {
        Dictionary<string, double> tableIdentifier = new Dictionary<string, double>();

        public override double VisitCompileUnit(CombinedParser.CompileUnitContext context)
        {
            return Visit(context.expression());
        }

        public override double VisitNumberExpr(CombinedParser.NumberExprContext context)
        {
            var result = double.Parse(context.GetText());
            Debug.WriteLine(result);

            return result;
        }

        //IdentifierExpr
        public override double VisitIdentifierExpr(CombinedParser.IdentifierExprContext context)
        {
            var result = context.GetText();
            double value;
            //видобути значення змінної з таблиці
            if (tableIdentifier.TryGetValue(result.ToString(), out value))
            {
                return value;
            }
            else
            {
                return 0.0;
            }
        }

        public override double VisitParenthesizedExpr(CombinedParser.ParenthesizedExprContext context)
        {
            return Visit(context.expression());
        }

        public override double VisitExponentialExpr(CombinedParser.ExponentialExprContext context)
        {
            var left = WalkLeft(context);
            var right = WalkRight(context);

            Debug.WriteLine("{0} ^ {1}", left, right);
            return System.Math.Pow(left, right);
        }//

        public override double VisitAdditiveExpr(CombinedParser.AdditiveExprContext context)
        {
            var left = WalkLeft(context);
            var right = WalkRight(context);

            if (context.operatorToken.Type == CombinedParser.ADD)
            {
                Debug.WriteLine("{0} + {1}", left, right);
                return left + right;
            }
            else //LabCalculatorLexer.SUBTRACT
            {
                Debug.WriteLine("{0} - {1}", left, right);
                return left - right;
            }
        }

        public override double VisitMultiplicativeExpr(CombinedParser.MultiplicativeExprContext context)
        {
            var left = WalkLeft(context);
            var right = WalkRight(context);

            if (context.operatorToken.Type == CombinedLexer.MULTIPLY)
            {
                Debug.WriteLine("{0} * {1}", left, right);
                return left * right;
            }
            else //LabCalculatorLexer.DIVIDE
            {
                Debug.WriteLine("{0} / {1}", left, right);
                return left / right;
            }
        }

        public override double VisitMinMaxExpression([NotNull] CombinedParser.MinMaxExpressionContext context)
        {
            double vl = Visit(context.GetRuleContext<CombinedParser.ExpressionContext>(0));
            double vr = Visit(context.GetRuleContext<CombinedParser.ExpressionContext>(1));
            if(context.operatorToken.Type == CombinedLexer.MIN)
            {
                return (vl <= vr ? vl : vr);
            }
            else //MAX
            {
                return (vl <= vr ? vr : vl);
            }
        }

        private double WalkLeft(CombinedParser.ExpressionContext context)
        {
            return Visit(context.GetRuleContext<CombinedParser.ExpressionContext>(0));
        }

        private double WalkRight(CombinedParser.ExpressionContext context)
        {
            return Visit(context.GetRuleContext<CombinedParser.ExpressionContext>(1));
        }
    }
}
