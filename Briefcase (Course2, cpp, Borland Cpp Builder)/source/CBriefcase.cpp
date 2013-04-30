#include "CBriefcase.h"

/**********************************************************************
***********************************************************************

                       Вычисление времени работы

***********************************************************************
***********************************************************************/

void _setTime()
{
	DeltaTime = GetTickCount();
}

void _getTime()
{
	DeltaTime = GetTickCount()-DeltaTime; 
}

/**********************************************************************
***********************************************************************

                             CBriefcase

***********************************************************************
***********************************************************************/

// ------------------ CBriefcase ----------------------
CBriefcase::CBriefcase(int count, int maxcost)
{
    invest = NULL;
    cur = NULL;
    best = NULL;
    NumItem = 0;
    MaxNumItem = count;
    MaxCost = maxcost;

    if (count)
    {
        invest = new SInvestment[count];
        cur = new bool[count];
        best = new bool[count];
        for (int i=0; i<count; i++)
            ZeroMemory((void*)&invest[i], sizeof(SInvestment));
    }
}

// ------------------ CBriefcase copy ----------------------
CBriefcase::CBriefcase(const CBriefcase &v)
{
    invest = NULL;
    cur = NULL;
    best = NULL;
    NumItem = 0;
    MaxNumItem = 0;
    
    copyfrom(v);
}

// ------------------ ~CBriefcase ----------------------
CBriefcase::~CBriefcase()
{
    clear();
}

// ------------------ copyfrom ----------------------
void CBriefcase::copyfrom(const CBriefcase &v)
{
    clear();

    NumItem = v.NumItem;
    MaxNumItem = v.MaxNumItem;
    MaxCost = v.MaxCost;
    copy(v.invest, invest, false);
    cur = new bool[MaxNumItem];
    best = new bool[MaxNumItem];
}

// ------------------ clear ----------------------
void CBriefcase::clear()
{
    if (!NumItem) return;
    
    for (int i=0; i<NumItem; i++)
    {
        delete[] invest[i].name;
    }
    
    delete[] invest;
    invest = NULL;
    delete[] best;
    best = NULL;
    delete[] cur;
    cur = NULL;

    MaxNumItem = 0;
    NumItem = 0;
}

// ------------------ copy ----------------------
void CBriefcase::copy(SInvestment * &s, SInvestment * &d, bool del, BYTE inc, bool create)
{
    if (create)
        d = new SInvestment[NumItem+inc];
    
    for (int i=0; i<(NumItem); i++)
    {
        d[i].name = new char[strlen(s[i].name)+1];
        strcpy(d[i].name, s[i].name);
        if (del)
            delete[] s[i].name;
        d[i].cost = s[i].cost;
        d[i].profit = s[i].profit;
    }
    if (del)
    {
        delete[] s;
        s = NULL;
    }
}

// ------------------ copybool ----------------------
void CBriefcase::copybool(bool * &s, bool * &d, BYTE inc, bool create)
{
    if (create)
        d = new bool[NumItem+inc];

    for (int i=0; i<NumItem; i++)
        d[i] = s[i];
}

// ------------------ resize ----------------------
void CBriefcase::resize()
{
    SInvestment *tmp=NULL;

    copy(invest, tmp, true);
    copy(tmp, invest, true, 5);

    bool *tmp2;

    copybool(best, tmp2);
    delete[] best;
    copybool(tmp2, best, 5);

    copybool(cur, tmp2, false);
    delete[] cur;
    copybool(tmp2, cur, 5);

    MaxNumItem += 5;

    ZeroMemory(&invest[NumItem], sizeof(SInvestment));
}

// ------------------ add ----------------------
bool CBriefcase::add(char *a, int c, int p)
{
    for (int i=0; i<NumItem; i++)
        if (!strcmp(invest[i].name, a)) return false;

    if (NumItem == MaxNumItem) resize();

    invest[NumItem].name = new char[strlen(a)+1];
    strcpy(invest[NumItem].name, a);
    invest[NumItem].cost = c;
    invest[NumItem].profit = p;

    NumItem++;

    return true;
}

// ------------------ del ----------------------
bool CBriefcase::del(char *n)
{
	for (int i=0; i<NumItem; i++)
		if (!strcmp(invest[i].name, n))
        {
        	if (i < NumItem-1)
            {
	        	delete[] invest[i].name;                
    	        int len_ = (NumItem-i-1)*sizeof(SInvestment);
        	    memcpy(&invest[i], &invest[i+1], len_);
            }
            NumItem--;
		    invest[NumItem].name = NULL;
            return true;            
        }

    return false;

}

// ------------------ del ----------------------
bool CBriefcase::del(unsigned int num)
{
	if (num >= NumItem) return false;

    delete[] invest[num].name;
    int len_ = (NumItem-num-1)*sizeof(SInvestment);
    memcpy(&invest[num], &invest[num+1], len_);

    NumItem--;
    invest[NumItem].name = NULL;
    return true;
}

// ------------------ setMaxCost ----------------------
void CBriefcase::setMaxCost(int maxcost)
{
    MaxCost = maxcost;
}

// ------------------ getMaxCost ----------------------
int CBriefcase::getMaxCost()
{
    return MaxCost;
}

// ------------------ getNumItem ----------------------
int CBriefcase::getNumItem()
{
    return NumItem;
}

// ------------------ calcBestVariant ----------------------
void CBriefcase::calcBestVariant()
{
    _setTime();

    for (int i=1; i<=ITER_COUNT; i++)
    {
	    NoUsedProfit = 0;
    	BestCost = 0;
	    BestProfit = 0;
    	CurCost = 0;
	    CurProfit = 0;

    	for (int i=0; i<(int)NumItem; i++)
	    {
    	    NoUsedProfit += invest[i].profit;
        	cur[i] = false;
	        best[i] = false;
    	}

	    search();
	}

    _getTime();
}

// ------------------ calcBestVariant ----------------------
bool CBriefcase::exists(char *n)
{
    for (int i=0; i<(int)NumItem; i++)
    	if (!strcmp(invest[i].name, n))
        	return true;

    return false;
}

/**********************************************************************
***********************************************************************

                             CBrunchAndBound
                             
***********************************************************************
***********************************************************************/

// ------------------ CBrunchAndBound ----------------------
CBrunchAndBound::CBrunchAndBound(int count, int maxcost):CBriefcase(count, maxcost) {}

// ------------------ CBrunchAndBound copy ----------------------
CBrunchAndBound::CBrunchAndBound(const CBrunchAndBound &v):CBriefcase(v) {}

// ------------------ _search ----------------------
void CBrunchAndBound::_search(int ind)
{
    // это лист
    if (ind == NumItem)
    {
        copybool(cur, best, false);
        BestProfit = CurProfit;
        BestCost = CurCost;
        return;
    }

    // есть ли смысл расматривать поддерево, включающее [ind]
    if ((CurCost + invest[ind].cost) <= MaxCost)
    {
        cur[ind] = true;
        CurCost += invest[ind].cost;
        CurProfit += invest[ind].profit;
        NoUsedProfit -= invest[ind].profit;

        _search(ind+1);

        cur[ind] = false;
        CurCost -= invest[ind].cost;
        CurProfit -= invest[ind].profit;
        NoUsedProfit += invest[ind].profit;
    }

    // а если я выкину [ind], то будет ли его правое поддерево подходить по условию?
    NoUsedProfit -= invest[ind].profit;
    if ((CurProfit + NoUsedProfit) > BestProfit)
        _search(ind+1);
    NoUsedProfit += invest[ind].profit;
}

// ------------------ search ----------------------
void CBrunchAndBound::search()
{
	_search(0);
}


/**********************************************************************
***********************************************************************

                             CHillClimbing

***********************************************************************
***********************************************************************/

// ------------------ CHillClimbing ----------------------
CHillClimbing::CHillClimbing(int count, int maxcost):CBriefcase(count, maxcost) {}

// ------------------ CHillClimbing copy ----------------------
CHillClimbing::CHillClimbing(const CHillClimbing &v):CBriefcase(v) {}

// ------------------ search ----------------------
void CHillClimbing::search()
{
    int maxval, maxInd;

    for (int i=0; i<(int)NumItem; i++)
    {
        maxval = 0;
        maxInd = -1;
                
        for (int j=0; j<(int)NumItem; j++)
            if ( ( !best[j] ) && ( maxval < (int)invest[j].profit ) && ( (BestCost + invest[j].cost) <= MaxCost ) )
            {
                maxval = invest[j].profit;
                maxInd = j;
            }

        if (maxInd<0) break;

        BestCost += invest[maxInd].cost;
        best[maxInd] = true;
        BestProfit += invest[maxInd].profit; 

    }    
}


/**********************************************************************
***********************************************************************

                             CRandomSearch

***********************************************************************
***********************************************************************/

// ------------------ CRandomSearch ----------------------
CRandomSearch::CRandomSearch(int count, int maxcost):CBriefcase(count, maxcost) { randomize(); }

// ------------------ CRandomSearch copy ----------------------
CRandomSearch::CRandomSearch(const CRandomSearch &v):CBriefcase(v) {}

// ------------------ _Add ----------------------
bool CRandomSearch::_Add()
{
    // сколько элементов ещё можно добавить в решение, чтобы они уместились в пределах стоимости 
    int count = 0;

    for (int i=0; i<(int)NumItem; i++)
        if ((!cur[i]) && ((CurCost + invest[i].cost) <= MaxCost))
            ++count;

    if (!count) return false;

    // случайный выбор одного элемента
    int select = random(count);

    // получение случайного элемента
    int i = 0;

    for (; i<(int)NumItem; i++)
        if ((!cur[i]) && ((CurCost + invest[i].cost) <= MaxCost))
        {
            --select;
            if (select<1) break;
        }

    // добавление элемента в решение
    CurProfit += invest[i].profit;
    CurCost += invest[i].cost;
    cur[i] = true;

    return true;        
}

// ------------------ _Del ----------------------
bool CRandomSearch::_Del()
{
    // определение кол-ва элементов в решении
    int count = 0;

    for (int i=0; i<(int)NumItem; i++)
        if (cur[i]) ++count;

    if (!count) return false;

    // случайный выбор одного элемента
    int select = random(count);

    // получение случайного элемента
    int i = 0;

    for (; i<(int)NumItem; i++)
        if (cur[i])
        {
            --select;
            if (select<1) break;
        }

    // удаление элемента из решения
    CurProfit -= invest[i].profit;
    CurCost -= invest[i].cost;
    cur[i] = false;

    return true;    
}

/**********************************************************************
***********************************************************************

                             CSimulatedAnnealing

***********************************************************************
***********************************************************************/

// ------------------ CSimulatedAnnealing ----------------------
CSimulatedAnnealing::CSimulatedAnnealing(int count, int maxcost):CRandomSearch(count, maxcost)
{
    MaxUnchanged = 1;
    MaxSaved = 1;
    K_ch = 1;
}

// ------------------ CSimulatedAnnealing copy ----------------------
CSimulatedAnnealing::CSimulatedAnnealing(const CSimulatedAnnealing &v):CRandomSearch(v)
{
    MaxUnchanged = v.MaxUnchanged;
    MaxSaved = v.MaxSaved;
    K_ch = v.K_ch;
}

void CSimulatedAnnealing::setMaxUnchanged(int v)
{
    MaxUnchanged = v;
}

void CSimulatedAnnealing::setMaxSaved(int v)
{
    MaxSaved = v;
}

int CSimulatedAnnealing::getMaxUnchanged()
{
    return MaxUnchanged;
}

int CSimulatedAnnealing::getMaxSaved()
{
    return MaxSaved;
}

void CSimulatedAnnealing::setK_changed(int v)
{
    K_ch = v;
}

int CSimulatedAnnealing::getK_changed()
{
    return K_ch;
}

// ------------------ search ----------------------
void CSimulatedAnnealing::search()
{
    if (NumItem < 2) return;

    // поиск максимальной и минимальной прибыли
    int max_profit = invest[0].profit,
        min_profit = max_profit;

    for (int i=1; i<NumItem; i++)
    {
        if (max_profit < invest[i].profit)
            max_profit = invest[i].profit;
        if (min_profit > invest[i].profit)
            min_profit = invest[i].profit;
    }

    // начальное значение T
    float T = 0.75 * (max_profit - min_profit);

    // формирование начального случайного решения
    while (_Add());

    // принятия случайного решения как лучшего
    BestProfit = CurProfit;
    BestCost = CurCost;
    for (int i=0; i<NumItem; i++)
        best[i] = cur[i];

    int num_save = 0,
        num_unch = 0;

    bool save_changed,
         save;

    while (num_unch < MaxUnchanged)
    {
        // удаление K_ch случайных элементов
        for (int i=0; i<K_ch; i++)
            _Del();

        // добавление K_ch случайных элементов
        while (_Add());

        // если есть улучшение
        if (CurProfit > BestProfit)
        {
            save_changed = true;
            save = false;            
        } else
        if (CurProfit == BestProfit)
        {   // формула вероятности даст 1
            save = save_changed = false;
        } else
        {   // нужно ли сохранять решение
            save = save_changed = (rand() < expl(abs(CurProfit - BestProfit)/T));
        }

        // если нужно сохранить новое решение
        if (save_changed)
        {
            // сохраняю новое решение
            BestProfit = CurProfit;
            BestCost = CurCost;
            copybool(cur, best);
            num_unch = 0;   // было сохранено изменение
        } else
        {   // восстановление пердыдущего решения
            CurProfit = BestProfit;
            CurCost = BestCost;
            copybool(best, cur);
            ++num_unch;
        }

        // если сохранено решение, которое НЕ ЛУЧШЕ предыдущего
        if (save)
        {
            if (++num_save > MaxSaved)
            {
                num_save = 0;
                T *= 0.8;
                num_unch = 0;
            }
        } // if (save)

    } // while (num_unch < MaxUnchanged)

}


/**********************************************************************
***********************************************************************

                             CLocalOptimum

***********************************************************************
***********************************************************************/

// ------------------ CLocalOptimum ----------------------
CLocalOptimum::CLocalOptimum(int count, int maxcost):CRandomSearch(count, maxcost) {}

// ------------------ CLocalOptimum copy ----------------------
CLocalOptimum::CLocalOptimum(const CLocalOptimum &v):CRandomSearch(v) {}

// ------------------ search ----------------------
void CLocalOptimum::search()
{                     
    int bad_tests,		// кол-во неэффективных испытаний
        bad_changes;	// кол-во неэффективных изменений

    trial = new bool[NumItem];

    // Испытания поторяются до тех пор, пока
    // не произойдет max_bad_test испытаний без улучшений.
    bad_tests = 0;
    while (bad_tests <= max_bad_tests)
    {
        // формирование начального случайного решения
        while (_Add());

	    TrialProfit = CurProfit;
    	TrialCost = CurCost;
	    for (int i=0; i<NumItem; i++)
    	    trial[i] = cur[i];

        // Повторять до max_bad_changes отсутствия улучшений в решениях.
        bad_changes = 0;
        while (bad_changes <= max_bad_changes)
        {
            // Удаление num случайных элементов
            for (int i=0; i<num; i++)
            	_Del();

            // Добавление num случайных элементов
            while (_Add());

            // Если данное изменение улучшает решение, то сохраняю его.
            // Иначе восстанавливаю предыдущее.
            if (CurProfit > TrialProfit)
            {
                // Сохранение решения
                TrialProfit = CurProfit;
                TrialCost = CurCost;
                copybool(cur, trial);
                bad_changes = 0; // Сброс счетчика плохих изменений
            } else
            {
                // Восстановление решения
                CurProfit = TrialProfit;
                CurCost = TrialCost;
                copybool(trial, cur);
                ++bad_changes; // Очередное плохое изменение
            }
        } // while (bad_changes <= max_bad_changes)

        // Если это испытание лучше лучшего, то его нужно сохранить.
        if (TrialProfit > BestProfit)
        {
            BestProfit = TrialProfit;
            BestCost = TrialCost;
            copybool(trial, best);
            bad_tests = 0; // Сброс счетчика прохих решений
        } else
            ++bad_tests; // Очередное плохое решение

        // Очистка текущего решения
        CurProfit = 0;
        CurCost = 0;
        for (int i=0; i<NumItem; i++)
            cur[i] = false;
    }

    delete[] trial;
}

// ------------------ setNum ----------------------
void CLocalOptimum::setNum(int val)
{
    if (val>0)
		num = val;
}

// ------------------ setMaxBadTests ----------------------
void CLocalOptimum::setMaxBadTests(int val)
{
    if (val>0)
		max_bad_tests = val;
}

// ------------------ setMaxBadChanges ----------------------
void CLocalOptimum::setMaxBadChanges(int val)
{
    if (val>0)
		max_bad_changes = val;
}

// ------------------ getNum ----------------------
int CLocalOptimum::getNum()
{
	return num;
}

// ------------------ getMaxBadTests ----------------------
int CLocalOptimum::getMaxBadTests()
{
	return max_bad_tests;
}

// ------------------ getMaxBadChanges ----------------------
int CLocalOptimum::getMaxBadChanges()
{
	return max_bad_changes;
}

//######################################################################################################################
//######################################################################################################################
//######################################################################################################################

/**********************************************************************
***********************************************************************

                             CBalancedProfit

***********************************************************************
***********************************************************************/

//------------------- CBalancedProfit -------------
CBalancedProfit::CBalancedProfit(int count, int maxcost):CBriefcase(count, maxcost) {}

//------------------- CBalancedProfit copy --------
CBalancedProfit::CBalancedProfit(const CBalancedProfit &v):CBriefcase(v) {}

// ------------------ search ----------------------
void CBalancedProfit::search()
{
    int goodInd;
    float test_ratio=0.15, best_ratio;

    for (int i=0; i<(int)NumItem; i++)
    {
        best_ratio=0;
        goodInd = -1;
        for (int j=0; j<(int)NumItem; j++){
            test_ratio = float(invest[j].profit)/float(invest[j].cost);
            if ( ( !best[j] ) && ( best_ratio < test_ratio ) && ( (BestCost + invest[j].cost) <= MaxCost ) )
            {
                best_ratio = test_ratio;
                goodInd = j;
            }
        }

        if (goodInd<0) break;

        BestCost += invest[goodInd].cost;
        best[goodInd] = true;
        BestProfit += invest[goodInd].profit;
    }
}

/**********************************************************************
***********************************************************************

                             CRandomSearchMod

***********************************************************************
***********************************************************************/

//------------------- CRandomSearchMod ------------
CRandomSearchMod::CRandomSearchMod(int count, int maxcost) : CRandomSearch(count,maxcost) {}

//------------------- CRandomSearchMod copy -------
CRandomSearchMod::CRandomSearchMod(const CRandomSearchMod &v) : CRandomSearch(v) {}

//------------------- search ----------------------
void CRandomSearchMod::search(){
     int trial;

     //делаем несколько испытаний и сохраняем лучшее
     for (trial = 0; trial < NumItem; trial++){
        //Случайный выбор до тех пор, пока не исчерпан лимит средств
        while (_Add());

        //Лучше полученное решение предыдущего?
        if (CurProfit > BestProfit){
            BestProfit = CurProfit;
            BestCost = CurCost;
            for (int i = 0; i < NumItem; i++) best[i] = cur[i];
        }

        //Сброс исследуемого решения для следующего испутания
        CurProfit = 0;
        CurCost = 0;
        for (int i = 0; i < NumItem; i++) cur[i] = false;
     }
}

/**********************************************************************
***********************************************************************

                             CIncrementalImprovement

***********************************************************************
***********************************************************************/

//------------------- CIncrementalImprovement -----
CIncrementalImprovement::CIncrementalImprovement(int count, int maxcost) : CRandomSearch(count,maxcost) {}

//------------------- CIncrementalImprovement copy -
CIncrementalImprovement::CIncrementalImprovement(const CIncrementalImprovement &v) : CRandomSearch(v) {}

//------------------- search ----------------------
void CIncrementalImprovement::search(){
     int trial;
     //Генерируем случайное решение
     while (_Add());

     //Сохраняем найденное решение
     BestCost = CurCost;
     BestProfit = CurProfit;
     for (int i = 0; i < NumItem; i++) best[i] = cur[i];

     //Несколько раз удаляем произвольную инвестицию и добавляем новую(ые) и сохраняем лучшее
     for (trial = 0; trial < NumItem; trial++){
        //Удаляем произвольную инвестицию
        _Del();

        //Случайный выбор до тех пор, пока не исчерпан лимит средств
        while (_Add());

        //Лучше полученное решение предыдущего?
        //Если Да, то сохраняем изменения, если Нет, то отказываемся от полученного решения
        if (CurProfit > BestProfit){
            BestProfit = CurProfit;
            BestCost = CurCost;
            for (int i = 0; i < NumItem; i++) best[i] = cur[i];
        }
        else{
            CurProfit = BestProfit;
            CurCost = BestCost;
            for (int i = 0; i < NumItem; i++) cur[i] = best[i];
        }
     }
}

/**********************************************************************
***********************************************************************

                             CStopMoment

***********************************************************************
***********************************************************************/

//------------------- CStopMoment -----------------
CStopMoment::CStopMoment(int count, int maxcost) : CRandomSearch(count,maxcost) {}

//------------------- CStopMoment copy ------------
CStopMoment::CStopMoment(const CStopMoment &v) : CRandomSearch(v) {}

//------------------- search ----------------------
void CStopMoment::search(){
     int bad_trials, non_changes;
     bad_trials = 0;

     //Испытания повторяются до тех пор,
     //пока не досигнем NumItem испытаний в строке без изменений
     do{
        //Генерируем случайное решение
        while (_Add());

        //Начинаем работать с этим решением как с испытываемым
        int trial_cost = CurCost;
        int trial_profit = CurProfit;
        bool *trial_bool = new bool[NumItem];
        for (int i = 0; i < NumItem; i++) trial_bool[i] = cur[i];

        //Повторяем до тех пор, пока не попробуем
        //NumItem изменений в строке без улучшения
        non_changes = 0;
        while (non_changes < NumItem){
            //Удаление случайного элемента
            _Del();

            //Случайный выбор до тех пор, пока не исчерпан лимит средств
            while (_Add());

            //Если это улучшает решение, сохраняем его,
            //в противном случае востанавливаем исходные значения
            if (CurProfit > trial_profit){
                trial_profit = CurProfit;
                trial_cost = CurCost;
                for (int i = 0; i < NumItem; i++) trial_bool[i] = cur[i];
                non_changes = 0; //Это хооршее изменение
            }
            else{
                CurProfit = trial_profit;
                CurCost = trial_cost;
                for (int i = 0; i < NumItem; i++) cur[i] = trial_bool[i];
                ++non_changes; //Это плохое изменение
            }
        }  //Конец While - попытки изменения

        //Если испытание является наилучшим решением на данный момент,
        //сохраняем его
        if (trial_profit > BestProfit){
            BestProfit = trial_profit;
            BestCost = trial_cost;
            for (int i = 0; i < NumItem; i++) best[i] = trial_bool[i];
            bad_trials=0; // Это хорошее испытание
        }
        else ++bad_trials; //Это плохое испытание

        //Сброс исследуемого решения для следующего испытания
        CurProfit = 0;
        CurCost = 0;
        for (int i = 0; i < NumItem; i++) cur[i] = false;
     }
     while (bad_trials >= NumItem);
}
