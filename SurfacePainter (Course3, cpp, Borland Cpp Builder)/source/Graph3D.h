#ifndef Graph3DH
#define Graph3DH

#include <vcl.h>
#include <windows.h>
#include <math.h>

const FLOAT_16 = (16*sizeof(float));

// папаметры поверхности
const int SurfX = 60; // не ставь меньше 3,
const int SurfZ = 60; // будет вылетать 
const float stepX = 2.0/(SurfX-1);
const float stepZ = 2.0/(SurfZ-1);

struct FPOINT
{
	float x, y, z;
};

struct IPOINT
{
	int x, y;
};

// тип функции получения высот поверхности
typedef float(*SURF_FUNC)(float x, float z);

inline FPOINT FPoint(float x, float y, float z);
inline IPOINT IPoint(int x, int y);

#define CopyMatrix(s, d) CopyMemory((void*)(d), (void*)(s), FLOAT_16)

	class CGraph3D
	{
		protected:
				// сама поверхность
			FPOINT surface[SurfX][SurfZ];
			
				// текущая матрица вида
			float ViewMatrix[16];

				// половина ширины и высоты области вывода
			int W2, H2;

				// положение и направление взгляда
			FPOINT cam_pos, cam_ang;

            	// фокус камеры
			float focus;

				// на чем выводить
			TCanvas *Canvas;

				// проецирование точки из локальной системы координат на плоскость экрана
			void ProjectPoint(FPOINT *p1, int &x, int &y);
				// перевод ГСО -> ЛСО
			void Global2Local(FPOINT *s, FPOINT *d);
				// перемножение матриц r = m1*m2
			void MulMatrix(const float m1[16], const float m2[16], float r[16]);
				// матричный перенос
			void _Translate_(FPOINT &d);
				// матричный поворот
			void _Rotate_(FPOINT &d);
				// формирование единичной матрицы
			void MakeSingle(float m[16]);
				// перевод мировых координат точки сразу в экранные координаты
			void Point3D_To_ScreenCoords(const FPOINT *p, FPOINT *loc, int &x, int &y);
				// вывод горизонтальной линии
			void HLine(int y, int x1, int x2, DWORD color);
				// вычисление нормати к треугольнику
			void Normal(FPOINT *p1, FPOINT *p2, FPOINT *p3, FPOINT *n);
				// определение цвета треугольника в зависимости от угла между его нормалью и направлением взгляда
			DWORD TrColor(FPOINT *p1, FPOINT *p2, FPOINT *p3);
				// вывод треугольника
			void Triangle(IPOINT *p1, IPOINT *p2, IPOINT *p3, DWORD color);
				// подмена индексов массива поверхности для текущего направления взгляда
			void ChangeIndexies_Painter(int &x, int &z);

		public:
			CGraph3D(TCanvas *c);
			~CGraph3D();

				// сброс всех параметров в начальные
			void ReInit();
				// изменение размеров области вывода
			void ChangeBounds(int w, int h);

				// формирование поверхности
			void InitSurface(SURF_FUNC func);
				// вывод поверхности
			void DrawSurface();

				// перенос поверхности НА указанные ПРИРАЩЕНИЯ
			void Translate(float x, float y, float z);
				// поворот поверхности НА указанные ПРИРАЩЕНИЯ
			void Rotate(float x, float y, float z);
	};

#endif // Graph3DH
