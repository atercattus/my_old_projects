#ifndef BRIEFCASE_H
#define BRIEFCASE_H

    #include <vcl.h>
    #include <windows.h>
    #include <math.h>
    #include <mem.h>

	// больше не ругается "W8012 Comparing signed and unsigned values"
	#pragma warn -csu

	// больше не ругается "W8032 Temporary used for parameter '???' in call to '???'"
	#pragma warn -lvc

    // больше не ругается "W8004 'identifier' is assigned a value that is never used"
    #pragma warn -aus
    

    // СТРУКТУРА ИНВЕСТИЦИИ
    struct SInvestment
    {
        char *name;               // название инвестиции
        int cost;                 // стоимость
        int profit;               // прибыль
    };

    unsigned int ITER_COUNT=1;
    DWORD DeltaTime=0;

    #define STD_MAX_COST 100


/**************************************
           CBriefcase
***************************************/
    class CBriefcase
    {
        protected:
            SInvestment *invest;                      // инвестиции
            int NumItem;                              // число инвестиций в портфеле
            int MaxNumItem;                           // максимальное число инвестиций в портфеле
            int MaxCost;                              // размер портфеля

            int NoUsedProfit;                         // своюодная прибыль

            bool *best;                               // лучшее решение
            int BestCost;                             // параметры лучшего решения
            int BestProfit;

            bool *cur;                                // текущее решение
            int CurCost;                              // параметры текущего решения
            int CurProfit;

            bool *trial;                              // текущее решение в алгоритмах последовательного улучшения
            int TrialCost;                            // параметры текущего решения
            int TrialProfit;                          // для алгоритмов последовательного улучшения

            virtual void search()=0;                  // рекурсивное вычисление, по вариантам

            void copy(SInvestment * &s, SInvestment * &d,  bool del, BYTE inc=0, bool create=true); // копирование массивов инвестиций
            void copybool(bool * &s, bool * &d, BYTE inc=0, bool create=true);
            void resize();                            // увеличение размера массива на 1
        public:
            CBriefcase(int count=0, int maxcost=STD_MAX_COST); // конструктор с указанием кол-ва инвестиций и максимальной стоимости
            CBriefcase(const CBriefcase &v);          // конструктор копирования
            ~CBriefcase();

            bool add(char *a, int c, int p);          // добавление инвестиции
            bool del(char *n);						  // удаление инвестиции
            bool del(unsigned int num);				  // удаление инвестиции
            void clear();                             // очистка портфеля
            bool exists(char *n);					  // проверка на существование инвестиции с указанным именем

            void setMaxCost(int maxcost);             // установка максимальной стоимости
            int getMaxCost();
            int getNumItem();

            virtual void copyfrom(const CBriefcase &v); // копия объекта

            virtual void calcBestVariant();             // вычисление лучшего варианта, общее

            friend void outbestBriefcase(const CBriefcase &v);  // вывод содержимого портфеля
            friend void outBriefcase();                         // вывод списка инвестиций
            friend void saveBriefcase(char *path);              // сохранение и загрузка
            friend void loadBriefcase(char *path);              // списка инвестиций
    };

//######################################################################################################################


/**************************************
           CBrunchAndBound
***************************************/
    class CBrunchAndBound : public CBriefcase
    {
        protected:
            virtual void search();
            void _search(int);
        public:
            CBrunchAndBound(int count=0, int maxcost=STD_MAX_COST);
            CBrunchAndBound(const CBrunchAndBound &v);      
    };


/**************************************
           CHillClimbing
***************************************/
    class CHillClimbing : public CBriefcase
    {
        protected:
            virtual void search();
        public:
            CHillClimbing(int count=0, int maxcost=STD_MAX_COST);
            CHillClimbing(const CHillClimbing &v);        
    };
    

/**************************************
           CRandomSearch
***************************************/
    class CRandomSearch : public CBriefcase
    {
        protected:
            bool _Add();    // добавление случайного элемента к исследуемому решению
            bool _Del();    // удаление случайного элемента из исследуемого решения
        public:
            CRandomSearch(int count=0, int maxcost=STD_MAX_COST);
            CRandomSearch(const CRandomSearch &v);        
    };


/**************************************
           CSimulatedAnnealing
***************************************/
    class CSimulatedAnnealing : public CRandomSearch
    {
        protected:
            virtual void search();
            int MaxUnchanged;
            int MaxSaved;
            int K_ch;
        public:
            CSimulatedAnnealing(int count=0, int maxcost=STD_MAX_COST);
            CSimulatedAnnealing(const CSimulatedAnnealing &v);

            void setMaxUnchanged(int);
            int getMaxUnchanged();
            void setMaxSaved(int);
            int getMaxSaved();
            void setK_changed(int);
            int getK_changed();
    };
    

/**************************************
           CLocalOptimum
***************************************/
    class CLocalOptimum : public CRandomSearch
    {
        protected:
            virtual void search();

            /*
                Одновременное изменение num элементов для улучшения тестируемого решения.
                Испытания повторяются до тех пор, пока в течение max_nochanged_test
              испытаний не произойдет улучшений.
                В течение каждого испытания вносятся случайные изменения, пока в течение
              max_nobest_changes испытаний не будет улучшений.
            */

            int num;
            int max_bad_tests;
            int max_bad_changes;

        public:
            CLocalOptimum(int count=0, int maxcost=STD_MAX_COST);
            CLocalOptimum(const CLocalOptimum &v);

            void setNum(int);
            void setMaxBadTests(int);
            void setMaxBadChanges(int);
            int getNum();
            int getMaxBadTests();
            int getMaxBadChanges();
    };

//######################################################################################################################

/**************************************
           CBalancedProfit
**************************************/
    class CBalancedProfit : public CBriefcase
    {
        protected:
            virtual void search();
        public:
            CBalancedProfit(int count=0, int maxcost=STD_MAX_COST);
            CBalancedProfit(const CBalancedProfit &v);
    };


/**************************************
           CRandomSearchMod
**************************************/
    class CRandomSearchMod : public CRandomSearch
    {
        protected:
            virtual void search();
        public:
            CRandomSearchMod(int count=0, int maxcost=STD_MAX_COST);
            CRandomSearchMod(const CRandomSearchMod &v);
    };


/**************************************
           CIncrementalImprovement
**************************************/
    class CIncrementalImprovement : public CRandomSearch
    {
        protected:
            virtual void search();
        public:
            CIncrementalImprovement(int count=0, int maxcost=STD_MAX_COST);
            CIncrementalImprovement(const CIncrementalImprovement &v);
    };


/**************************************
               CStopMoment
**************************************/
    class CStopMoment : public CRandomSearch
    {
        protected:
            virtual void search();
        public:
            CStopMoment(int count=0, int maxcost=STD_MAX_COST);
            CStopMoment(const CStopMoment &v);
    };



#endif // BRIEFCASE_H
