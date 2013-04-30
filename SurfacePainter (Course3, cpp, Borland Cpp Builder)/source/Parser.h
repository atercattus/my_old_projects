#ifndef ParserH
#define ParserH

#include <fstream.h>
#include <windows.h>
#include <ctype.h>
#include <math.h>

    // возможные токены строки разбора
	enum EToken
	{
        tNothing = -1,

		tBeg,		// "("
		tEnd,		// ")", никогда не находится в стеке операторов

		tX,			// "x"
		tZ,			// "z"

		tConst,		// FPC

		tAdd,		// +
		tSub,		// -
		tMul,		// *
		tDiv,		// /
		tMod,		// %

		tSin,		// sin
		tCos,		// cos
		tTan		// tan
	};

    // структура-описатель токена
	struct SInfo
	{
		// тип
		EToken	token;
		// указатель на собственные данные
		float	data;
	};

	SInfo Info(EToken tok, float f);

	// один элемент строки разбора
	struct SItem
	{
		SInfo	info;
		SItem	*next;
	};

// #### Определение приоритета токена ####
	int Prior(EToken tok);

// #### Определение кол-ва значений, необходимых для оператора в стеке операндов ####
// { Необходимо для сортировки стека операторов с учетом их приоритетов }   
	int VarCount(EToken tok);


// #### Стек (без комментариев :) ####
	class CStack
	{
		private:
			SItem	*top;

		public:
			CStack();
			CStack(const CStack &v);
			~CStack();
			void push(SInfo i);
			void add_last(SInfo i);
			SInfo pop();
			SInfo read();
			bool empty();

            friend class CParser;
	};


// #### Разбор и вычисление выражения ####
	class CParser
	{
		private:
			char str[100];
			int len;

			// временные стеки операндов и операторов
			CStack	*vars;
			CStack	*ops;

            // результат разбора строки
			CStack	variables;
			CStack	operators;

			// при считывании getNextToken константы, здесь сохраняется ее значение
			float	data;

			// получение следующего токена из строки разбора
			EToken getNextToken(char *buf, int &pos);

			// выполнение одного блока операторов
			// если в стеке ops: "+(*+(/--+", то выполнится только "/--+"
			// и последний "(" удалится из стека
			bool callStacks(CStack *op, CStack *var);

			// копирование содержимого 'variables' и 'operators' в
			// 'vars' и 'ops' соответственно с учетом значений 'X' и 'Z'
			void InitNewEval(float x, float z);

			// удаление из памяти vars и ops
			void DoneCurEval();

            // внутренняя вычислялка значения
			bool Calc();

            // сортировка операторов и соответствующих операндов для учета приоритетов
			void SortBlock(CStack *op, CStack *var);

		public:
			CParser();
			~CParser();

			// загрузка формулы
			bool Parse(char *path);
			// вычисление Y по заданным X,Z
			bool Eval(float x, float z, float &y);
	};

#endif // ParserH
