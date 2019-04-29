grammar Combined;

/*
 * Parser Rules
 */

compileUnit : expression EOF;

expression :
	LPAREN expression RPAREN #ParenthesizedExpr
	| expression EXPONENT expression #ExponentialExpr
    | expression operatorToken=(MULTIPLY | DIVIDE) expression #MultiplicativeExpr
	| expression operatorToken=(ADD | SUBTRACT) expression #AdditiveExpr
	| operatorToken = (MIN | MAX) LPAREN expression COMA expression RPAREN #MinMaxExpression 
	| NUMBER #NumberExpr
	| IDENTIFIER #IdentifierExpr
	; 

/*
 * Lexer Rules
 */

NUMBER : INT (',' INT)?; //change to ','
IDENTIFIER : [a-zA-Z]+[1-9][0-9]*;

INT : ('0'..'9')+;

EXPONENT : '^';
MULTIPLY : '*';
DIVIDE : '/';
SUBTRACT : '-';
ADD : '+';
LPAREN : '(';
RPAREN : ')';
COMA : ';';
MIN : 'Min';
MAX : 'Max';

WS : [ \t\r\n] -> channel(HIDDEN);

