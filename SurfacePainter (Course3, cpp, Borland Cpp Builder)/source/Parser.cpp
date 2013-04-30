#include "Parser.h"

//######################################################################################################################
int Prior(EToken tok)
{
	switch (tok)
	{
		case tAdd	: return 1;
		case tSub	: return 1;
		case tMul	: return 2;
		case tDiv	: return 2;
		case tMod	: return 2;

		case tSin	: return 3; // ?
		case tCos	: return 3; // ?
		case tTan	: return 3; // ?

		default		: return 0; // конец строки всегда "выполн€етс€ последним"
	}
}

//######################################################################################################################
int VarCount(EToken tok)
{
	switch (tok)
	{
		case tAdd	: return 2;
		case tSub	: return 2;
		case tMul	: return 2;
		case tDiv	: return 2;
		case tMod	: return 2;

		case tSin	: return 1;
		case tCos	: return 1;
		case tTan	: return 1;

		default		: return 0;
	}
}

//######################################################################################################################
SInfo Info(EToken tok, float f)
{
	SInfo info;
	info.token = tok;
	info.data = f;

	return info;
}

//######################################################################################################################
CStack::CStack()
{
	top = NULL;
}

//----------------------------------------------------------------------------------------------------------------------
CStack::CStack(const CStack &v)
{
	SItem *cur = v.top;

	top = NULL;

	while (cur)
	{
		add_last(cur->info);
		cur = cur->next;
	}
}

//----------------------------------------------------------------------------------------------------------------------
CStack::~CStack()
{
	while (!empty()) pop();
}

//----------------------------------------------------------------------------------------------------------------------
void CStack::push(SInfo i)
{
	SItem *cur = new SItem;
	cur->info = i;
	cur->next = top;
	top = cur;
}

//----------------------------------------------------------------------------------------------------------------------
void CStack::add_last(SInfo i)
{
	SItem *cur = top, *prev=NULL;
	while (cur)
	{
		prev = cur;
		cur = cur->next;
	}

	cur = new SItem;
	cur->info = i;
	cur->next = NULL;

	((prev) ? (prev->next) : top) = cur;
}

//----------------------------------------------------------------------------------------------------------------------
SInfo CStack::pop()
{
	if (empty()) return Info(tNothing, 0);

	SInfo val = top->info;
	SItem *cur = top;
	top = top->next;
	delete cur;

	return val;
}


//----------------------------------------------------------------------------------------------------------------------
SInfo CStack::read()
{
	if (!empty()) return top->info;

	return Info(tNothing, 0);
}

//----------------------------------------------------------------------------------------------------------------------
bool CStack::empty()
{
	return (top==NULL); // так корректнее, чем просто "return top;"
}

//######################################################################################################################
CParser::CParser()
{
	vars = NULL;
	ops = NULL;
}

//----------------------------------------------------------------------------------------------------------------------
CParser::~CParser()
{
	DoneCurEval();
}

//----------------------------------------------------------------------------------------------------------------------
EToken CParser::getNextToken(char *buf, int &pos)
{
	char ch;

	while ( (ch = buf[pos]) == ' ') pos++;

	if (isalpha(ch)) // с начала символ: 'x', 'z', 'sin', 'cos', 'tan'
	{
		switch (ch) // глупо, но пон€тно ¶)
		{
			case 'x':
			{
				pos++;
				return tX;
			}

			case 'z':
			{
            	pos++;
				return tZ;
			}

			default: // а остальные допустимые 3х-символьные (+PI, хочешь удал€й)
			{
				char tmp[4];
				strncpy(tmp, &buf[pos], 3);
				tmp[3] = 0; // ?

				if (isalpha(buf[pos+3])) return tNothing;

				if (!strcmp(tmp, "sin"))
				{
					pos += 3;
					return tSin;
				}
				if (!strcmp(tmp, "cos"))
				{
					pos += 3;
					return tCos;
				}
				if (!strcmp(tmp, "tan"))
				{
					pos += 3;
					return tTan;
				}

				// теперь пытаюсь сопоставить с PI

				if (isalpha(buf[pos+2])) return tNothing;

				if (!strncmp(tmp, "pi", 2))
				{
					pos += 2;
					data = M_PI;
					return tConst;
				}

			} // default
		} // switch

		return tNothing; // ошибка

	}
	else
	if (isdigit(ch)) // с начала цифра: только число
	{
		char *endptr;
		data = strtod(&buf[pos], &endptr);
		pos += (endptr-&buf[pos]);

		return tConst;

	}
	else
	{
		pos++;

		switch (ch)
		{
			case '+': return tAdd;
			case '-': return tSub;
			case '*': return tMul;
			case '/': return tDiv;
			case '%': return tMod;

			case '(': return tBeg;
			case ')': return tEnd;

			default : return tNothing;
		}
	}

    return tNothing;
}

//----------------------------------------------------------------------------------------------------------------------
bool CParser::callStacks(CStack *op, CStack *var)
{
	float op1, op2;
	SInfo info;
	
	do {
		info = op->pop();

		switch (info.token) {
			case -1:
				return true; // стек операторов пуст

			case tBeg:
				return true; // дошел до конца блока

			case tAdd:
				op1 = var->pop().data;
				op2 = var->pop().data;
				var->push(Info(tConst, op1+op2));
				break;

			case tSub:
				op1 = var->pop().data;
				op2 = var->pop().data;
				var->push(Info(tConst, op2-op1));
				break;

			case tMul:
				op1 = var->pop().data;
				op2 = var->pop().data;
				var->push(Info(tConst, op2*op1));
				break;

			case tDiv:
				op1 = var->pop().data;
				op2 = var->pop().data;
				if (!op1) return false; // деление на 0
				var->push(Info(tConst, op2/op1));
				break;

			case tMod:
				op1 = var->pop().data;
				op2 = var->pop().data;
				if (!op1) return false; // деление на 0
				var->push(Info(tConst, int(op2)%int(op1)));
				break;

			case tSin:
				op1 = var->pop().data;
				var->push(Info(tConst, sin(op1)));
				break;

			case tCos:
				op1 = var->pop().data;
				var->push(Info(tConst, cos(op1)));
				break;

			case tTan:
				op1 = var->pop().data;
				var->push(Info(tConst, tan(op1)));
				break;

			default:
				return false; // ошибка

		} // switch

	} while (1);
}

//----------------------------------------------------------------------------------------------------------------------
void CParser::InitNewEval(float x, float z)
{
	DoneCurEval();

	// VARIABLES
	vars = new CStack();

	SItem *cur = variables.top;
	while (cur)
	{
		if (cur->info.token == tX)
		{
			vars->push(Info(tConst, x));
		}
		else
		if (cur->info.token == tZ)
		{
			vars->push(Info(tConst, z));
		}
		else
			vars->push(cur->info);

		cur = cur->next;

	} // while (cur)


	// OPERATORS
	ops = new CStack();

	cur = operators.top;
	while (cur)
	{
		ops->push(cur->info);
		cur = cur->next;
	}

}

//----------------------------------------------------------------------------------------------------------------------
void CParser::DoneCurEval()
{
	if (vars) delete vars;
	vars = NULL;
	if (ops) delete ops;
	ops = NULL;
}

//----------------------------------------------------------------------------------------------------------------------
bool CParser::Calc()
{ // специально расписал все варианты отдельными проверками
  // сократить код можно, но так, по-моему, пон€тнее 
	int pos = 0;
	EToken tok;

	while (pos<len)
	{
		tok = getNextToken(str, pos);
		if (tok == tNothing) return false; // конец строки разбора

		if (tok >= tAdd) // получил оператор
		{
			operators.push(Info(tok, 0));
			continue;
		}

		if (tok == tConst) // получил константу
		{
			variables.push(Info(tok, data));
			continue;
		}

		if ((tok == tBeg) || (tok == tEnd)) // начало или конец блока
		{
			if (tok == tEnd)
			{
				SortBlock(&operators, &variables);
				variables.push(Info(tok, 0)); // что б не паритmс€ с вычислением позиции в стеке операндов
			}

			operators.push(Info(tok, 0));

			continue;
		}

		if ((tok == tX) || (tok == tZ)) // переменные
		{
			variables.push(Info(tok, 0));
		}

	} // while
}

//----------------------------------------------------------------------------------------------------------------------
void swap_tokens(SItem *t1, SItem *t2)
{ // обмен токенов

	SInfo i = t1->info;
	t1->info = t2->info;
	t2->info = i;
}

SItem* PtrPlusNum(SItem *ptr, int num)
{ // получение указател€, равного "ptr += num"

	if (!ptr || (num<0)) return NULL;

	SItem *c = ptr;
	while (num-- && c) c = c->next;

	return c;
}

void move(SItem *from, int count, SItem *to, SItem *&top)
{ // перемещение блока токенов в стеке из одного места в другое

	if (!from || (count<=0)) return;

	SItem *last = PtrPlusNum(from, count);

	SItem *first = from->next;
	from->next = last->next;

	if (to)
	{
		to->next = first;
		last->next = to->next;
	}
	else
	{
		last->next = top;
		top = first;
	}
}

#define PRIOR(i) Prior((i)->info.token)
#define TOKEN(i) (i)->info.token

void CParser::SortBlock(CStack *op, CStack *var)
{
	SItem *ptrOp, *ptrVar, *ptrVarPrev;
	bool moved;

	if (!op) return;

	do
	{
		moved = false;
		ptrOp = op->top; ptrVar = var->top; ptrVarPrev = NULL;

		while (ptrOp->next) // до конца стека
		{
			if (TOKEN(ptrOp) == tBeg) break; // конец блока

		/* 1 */
			if (PRIOR(ptrOp) >= PRIOR(ptrOp->next))
			{
				ptrOp = ptrOp->next;
				ptrVarPrev = ptrVar;
				ptrVar = PtrPlusNum(ptrOp, VarCount(TOKEN(ptrOp)));
				continue;
			}

		/* 2 */
			swap_tokens(ptrOp, ptrOp->next);

		/* 3 */
			int pos = VarCount(TOKEN(ptrOp))-1;
			pos = max(pos, 1)-1;
			move(	PtrPlusNum(ptrVar, pos),		// начало
					VarCount(TOKEN(ptrOp->next)),	// длина
					ptrVarPrev,					 	// куда
					var->top );

		/* 4 */
			moved = true;
			break;

		} // while (ptrOp->next)

	} while (moved); // пока не будет перестановок
}

//----------------------------------------------------------------------------------------------------------------------
bool CParser::Parse(char *path)
{
	ifstream f(path, ios::in);
	if (!f) return false;

	f.getline(str, 100);
	f.close();

	len = strlen(str);

	MessageBox(NULL, str, "‘ормула поверхности", MB_OK | MB_ICONINFORMATION);

	return Calc();
}

//----------------------------------------------------------------------------------------------------------------------
bool CParser::Eval(float x, float z, float &y)
{
    // конкретизирую стеки под текущие значени€ точки
	InitNewEval(x, z);

	CStack var_cut;
	CStack op_cut;

	SItem *cur_var = vars->top;
	SItem *cur_op = ops->top;

	while (cur_op)
	{
		if (cur_op->info.token == tEnd)
		{
			// завершен очередной блок, вычисл€ю его
			if (!callStacks(&op_cut, &var_cut))
			{
				MessageBox(NULL, "ќшибка вычислени€", NULL, MB_OK | MB_ICONERROR);
				DoneCurEval();
				return false;
			}
			if (cur_var) cur_var = cur_var->next;
			cur_op = cur_op->next;
			continue;
		}

		if (cur_var && (cur_var->info.token != tEnd)) var_cut.push(cur_var->info);
		op_cut.push(cur_op->info);

		cur_op = cur_op->next;
		if (cur_var && (cur_var->info.token != tEnd)) cur_var = cur_var->next;
	}

    // если остались необработанные ќѕ≈–ј“ќ–џ, вычисл€ю значение стеков
	if (!op_cut.empty())
	{
		if (!callStacks(&op_cut, &var_cut))
		{
			MessageBox(NULL, "ќшибка вычислени€", NULL, MB_OK | MB_ICONERROR);
			DoneCurEval();
			return false;
		}
	}

	// возможно, что формула была задана единственной константой (e.g. "1.25") и ничего не вычисл€лось,
	// тогда просто копирую значение этой константы в результат
	if (!var_cut.empty())	y = var_cut.top->info.data;
	else					y = vars->top->info.data;

	DoneCurEval();

	return true;
}
