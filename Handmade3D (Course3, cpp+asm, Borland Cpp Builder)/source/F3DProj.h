/*******************************************************************
*                           Fast3DLibrary
*   Графическая библиотека для работы с 2D/3D графикой.
*   Не используются OpenGL, DirectX, Glide, SDL, ets.
*   Основана на Win32API. Для Borland C++ Builder 6.
*
*                            F3DProj.*
*   Модуль, ответственный за базовые операции пространственной
*   математики (операции с матрицами, проецирование, ets.)
*******************************************************************/
#ifndef F3DProjH
#define F3DProjH

#include <windows.h>
#include "F3DL.h"

// фокусное расстояние камеры
#define camera_focus 5

const MAT_WIDTH		= 4;
const MAT_HEIGHT	= 4;
const MAT_SIZE		= MAT_WIDTH * MAT_HEIGHT;

//###########################################################################################################
//####################################### Базовая математика ################################################

		// задание единичной матрицы
	U32 f3d_MakeSingle(double m[MAT_SIZE]);

		// копирование матрицы
	U32 f3d_CopyMatrix(const double s[MAT_SIZE], double d[MAT_SIZE]);

		// перемножение матриц: AB = A*B
	U32 f3d_MulMatrix(const double A[MAT_SIZE], const double B[MAT_SIZE], double AB[MAT_SIZE]);

		// проецирование точки P в точку R (R.z == P.z)
	U32 f3d_Project(const TRIPLE *P, TRIPLE *R);

		// преобразование координат точки P из глобальной систему координат (С.К.) в локальную R на основе матрицы вида VM
	U32 f3d_Localize(const TRIPLE *P, const double VM[MAT_SIZE], TRIPLE *R);

    	// перенос С.К. VM в новую С.К. VMNew
	U32 f3d_Translate(const double VM[MAT_SIZE], double VMNew[MAT_SIZE], TRIPLE d);

    	// масштабирование С.К. VM в С.К. VMNew с коеффициентами d
	U32 f3d_Scale(const double VM[MAT_SIZE], double VMNew[MAT_SIZE], TRIPLE d);

    	// поворот С.К. VM на углы d в новую С.К. VMNew
	U32 f3d_Rotate(const double VM[MAT_SIZE], double VMNew[MAT_SIZE], TRIPLE d);

    	// вычисление нормали через векторное произведение векторов
	U32 f3d_Normal(const TRIPLE *a, const TRIPLE *b, bool normalize, TRIPLE *n);

    	// вычисление угла между двумя векторами
	U32 f3d_Angle(const TRIPLE *a, const TRIPLE *b, double *angle);

#endif // F3DProjH
