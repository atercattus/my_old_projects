#include "Graph3D.h"

//######################################################################################################################
inline FPOINT FPoint(float x, float y, float z)
{
	FPOINT p;
	p.x = x; p.y = y; p.z = z;
	return p;
}

//######################################################################################################################
inline IPOINT IPoint(int x, int y)
{
	IPOINT p;
	p.x = x; p.y = y;
	return p;
}

//######################################################################################################################
CGraph3D::CGraph3D(TCanvas *c)
{
	Canvas = c;
	ReInit();

	for (int z=0; z<SurfZ; z++)
		for (int x=0; x<SurfX; x++)
			surface[x][z] = FPoint(x, 0.0, z);
}

//######################################################################################################################
CGraph3D::~CGraph3D()
{
	// ...
}

//######################################################################################################################
void CGraph3D::ReInit()
{
	cam_pos = FPoint(0, 0, 0);
	cam_ang = FPoint(0, 0, 0);
	MakeSingle(ViewMatrix);
	focus = 10.0;
	W2 = 1; H2 = 1;
}

//######################################################################################################################
void CGraph3D::ChangeBounds(int w, int h)
{
	W2 = w >> 1;
	H2 = h >> 1;
	W2--;
	H2--;
}

//######################################################################################################################
void CGraph3D::ProjectPoint(FPOINT *p1, int &x, int &y)
{
	// получение оконных координат точки, заданной в локальных координатах
    // проекция перспективная (вроде как)

	float zd = p1->z;

	if (zd != -focus)	zd = focus / (zd + focus);
	else				zd = 0.0;

	x = W2 + (p1->x * W2 * zd);
	y = H2 + (p1->y * H2 * zd);
}

//######################################################################################################################
void CGraph3D::Global2Local(FPOINT *s, FPOINT *d)
{
	// перевод координат вектора из глобальной системы координат в локальную
	// умножением вектора на матрицу вида

	float vm_i;
	float pnt[4] = {s->x, s->y, s->z, 1};
	float *res = (float*)d;

	for (int i=0; i<4; i++)
	{
		vm_i = 0;

		for (int j=0; j<4; j++)
			vm_i += ( ViewMatrix[(i << 2)+j] * pnt[j] );

		if (i != 3) res[i] = vm_i;
		
	} // for
}

//######################################################################################################################
void CGraph3D::MulMatrix(const float m1[16], const float m2[16], float r[16])
{
	for (int i=0; i<4; i++)
	{
		for (int j=0; j<4; j++)
		{
			float Rij = 0;

			for (int k=0; k<4; k++)
				Rij += m1[(i<<2)+k]*m2[(k<<2)+j];

			r[(i<<2)+j] = Rij;

		} // for j

	} // for i
}

//######################################################################################################################
void CGraph3D::_Translate_(FPOINT &d)
{
	const
	float matrix[16] = {		1,		0,		0,		-d.x,
								0,		1,		0,		-d.y,
								0,		0,		1,		-d.z,
								0,		0,		0,		1
						};

	float tmp[16];

	MulMatrix(ViewMatrix, matrix, tmp);

	CopyMatrix(tmp, ViewMatrix);
}

//######################################################################################################################
void CGraph3D::_Rotate_(FPOINT &d)
{
	// NewVM = Ax*Ay*Az*VM

	float s_a, c_a, s_b, c_b, s_g, c_g;

	s_a = sin(d.x); c_a = cos(d.x);
	s_b = sin(d.y); c_b = cos(d.y);
	s_g = sin(d.z); c_g = cos(d.z);

	const
	float matrix[16] =	{

						c_b * c_g,							c_b * s_g,				  -s_b,			0,
			(s_a * s_b * c_g) - (c_a * s_g),	(s_a * s_b * s_g) + (c_a * c_g),	s_a * c_b,		0,
			(c_a * s_b * c_g) + (s_a * s_g),	(c_a * s_b * s_g) - (s_a * c_g),	c_a * c_b,		0,
							0,									0,						0,			1

						};

	float tmp[16];

	MulMatrix(ViewMatrix, matrix, tmp);

	CopyMatrix(tmp, ViewMatrix);
}

//######################################################################################################################
void CGraph3D::MakeSingle(float m[16])
{
	float single[16] = {	1, 0, 0, 0,
							0, 1, 0, 0,
							0, 0, 1, 0,
							0, 0, 0, 1
						};

	CopyMatrix(single, m);
}

//######################################################################################################################
void CGraph3D::InitSurface(SURF_FUNC func)
{
	for (int z=0; z<SurfZ; z++)
		for (int x=0; x<SurfX; x++)
		{
			float x_, z_;
			x_ = float(2*x)/(SurfX-1) - 1.0;
			z_ = float(2*z)/(SurfZ-1) - 1.0;
			// высоты поверхности определяет указанная функция
			surface[x][z].x = x_;
			surface[x][z].z = z_;
			surface[x][z].y = func(x_, z_);
		}
}

//######################################################################################################################
void CGraph3D::Point3D_To_ScreenCoords(const FPOINT *p, FPOINT *loc, int &x, int &y)
{
	// перехожу в локальную систему координат
	Global2Local((FPOINT*)p, loc);
	// получаю экранные координаты точки
	ProjectPoint(loc, x, y);
}

//######################################################################################################################
void CGraph3D::HLine(int y, int x1, int x2, DWORD color)
{
	if (x1 > x2) // а может шаблон сделать?
	{
		int tmp = x1;
		x1 = x2;
		x2 = tmp;
	}

	Canvas->Pen->Color = (TColor)color;
	Canvas->MoveTo(x1, y);
	Canvas->LineTo(x2, y);
}

//######################################################################################################################
void CGraph3D::Normal(FPOINT *p1, FPOINT *p2, FPOINT *p3, FPOINT *n)
{
	FPOINT a, b;
	a = FPoint(p2->x - p1->x, p2->y - p1->y, p2->z - p1->z);
	b = FPoint(p2->x - p3->x, p2->y - p3->y, p2->z - p3->z);

	n->x = (a.y * b.z) - (a.z * b.y);
	n->y = (a.z * b.x) - (a.x * b.z);
	n->z = (a.x * b.y) - (a.y * b.x);
}

//######################################################################################################################
DWORD CGraph3D::TrColor(FPOINT *p1, FPOINT *p2, FPOINT *p3)
{
	FPOINT n;
	// нормаль к треугольнику
	Normal(p1, p2, p3, &n);

	// вектор "вперед"
	FPOINT global = FPoint(0.0, 0.0, 0);
	FPOINT f;
	Global2Local(&global, &f);

	float len_n = sqrt(n.x*n.x + n.y*n.y + n.z*n.z);
	float len_f = sqrt(f.x*f.x + f.y*f.y + f.z*f.z);

	float cos_an = fabs((n.x*f.x + n.y*f.y + n.z*f.z) / (len_n*len_f));
	byte b = (cos_an*256);
	return (b << 16) | (b << 8) | b;
}

//######################################################################################################################
inline void swap(IPOINT &p1, IPOINT &p2)
{
	IPOINT p = p1; p1 = p2; p2 = p;
}

void CGraph3D::Triangle(IPOINT *p1, IPOINT *p2, IPOINT *p3, DWORD color)
{
#define FILL_MODE

#ifdef FILL_MODE

	// сортирую вершины
	if (p1->y > p2->y) swap(*p1, *p2);
	if (p2->y > p3->y) swap(*p2, *p3);
	if (p1->y > p2->y) swap(*p1, *p2);

	int x1, x2, sy;

	float k13, k12, k23;

	if ((p3->y - p1->y) == 0) return; // вырожденный треугольник
	k13 = (float)(p3->x - p1->x) / (float)(p3->y - p1->y);

	int sy_py, sy_py2; // == sy/*in begin*/ - p{1|2}.y;

	if (!(p2->y - p1->y)) goto out_2_part; // нет нижней половины, либо она не видна
	k12 = (float)(p2->x - p1->x) / (float)(p2->y - p1->y);

// Нижняя часть

	sy_py = 0;

	for (sy = p1->y; sy <= p2->y; sy++, sy_py++)
	{
		x1 = p1->x + (sy_py * k13);
		x2 = p1->x + (sy_py * k12);
		HLine(sy, x1, x2, color);
	}
out_2_part:

// Верхняя часть

	if (!(p3->y - p2->y)) return; // нет верхей половины
	k23 = (float)(p3->x - p2->x) / (float)(p3->y - p2->y);

	sy = p2->y+1;

	sy_py = sy - p1->y;
	sy_py2 = sy - (p2->y+1);

	for (; sy <= p3->y; sy++, sy_py++, sy_py2++)
	{
		x1 = p1->x + (sy_py * k13);
		x2 = p2->x + (sy_py2 * k23);
		HLine(sy, x1, x2, color);
	}

#else // FILL_MODE

	Canvas->Pen->Color = clWhite;
	Canvas->MoveTo(p1->x, p1->y);
	Canvas->LineTo(p2->x, p2->y);
	Canvas->LineTo(p3->x, p3->y);
	Canvas->LineTo(p1->x, p1->y);

#endif // FILL_MODE
}

//######################################################################################################################
void CGraph3D::ChangeIndexies_Painter(int &x, int &z)
{
 // подробнее по правилам выбора смотри в картинке

    // вектор направления взгляда
	FPOINT *f = (FPOINT*)&ViewMatrix[8];

// вид спереди ***********************************

	if (f->z >= 0.5)
	{
		if (f->x <= 0.0) // 1
		{
			z = SurfZ-z-1;
		}
		else			// 8
		{
			x = SurfX-x-1;
			z = SurfZ-z-1;
		}

		return;
	}

// вид сзади ***********************************

	if ((f->x >= 0) && (f->z <= -0.5)) // 5
	{
		x = SurfX-x-1;
		return;
	}

	if ((f->x < 0) && (f->z <= -0.5)) // 4
	{
		// 4 заранее корректно
		// т.к. вывод писался для этой области
		return;
	}

// вывод слева ***********************************

	if (f->x >= 0.5)
	{

		if ((f->z >= 0) && (f->z < 0.5)) // 7
		{
			int t = x; x = z; z = t;
			z = SurfZ-z-1;
			x = SurfX-x-1;
		}

		if ((f->z < 0) && (f->z > -0.5)) // 6
		{
			int t = x; x = z; z = t;
			x = SurfX-x-1;
		}

        return;
	}

// вывод справа ***********************************

	if (f->x <= -0.5)
	{

		if ((f->z >= 0) && (f->z < 0.5)) // 2
		{
			int t = x; x = z; z = t;
			z = SurfZ-z-1;
			return;
		}

		if ((f->z < 0) && (f->z > -0.5)) // 3
		{
			int t = x; x = z; z = t;
		}

		return;
	}

}

//######################################################################################################################
void CGraph3D::DrawSurface()
{
	FPOINT	pnt1, pnt2, pnt3, pnt4,
			loc1, loc2, loc3, loc4;
	int x1, y1, x2, y2, x3, y3, x4, y4;

	// очищаю окно
	Canvas->Brush->Color = clBlack;
	Canvas->Pen->Color = clBlack;
	Canvas->Rectangle(Canvas->ClipRect);

	// рисую поверхность
	Canvas->Pen->Color = clWhite;

	MakeSingle(ViewMatrix);
	_Translate_(cam_pos);
	_Rotate_(cam_ang);

	for (int z=0; z<SurfZ-1; z++)
	{
		for (int x=0; x<SurfX-1; x++)
		{
			x1 = x; y1 = z; ChangeIndexies_Painter(x1, y1);
			pnt1.x = surface[x1][y1].x;
			pnt1.z = surface[x1][y1].z;
			pnt1.y = -surface[x1][y1].y;
			Point3D_To_ScreenCoords(&pnt1, &loc1, x1, y1); // текущая точка

			x2 = x+1; y2 = z; ChangeIndexies_Painter(x2, y2);
			pnt2.x = surface[x2][y2].x;
			pnt2.z = surface[x2][y2].z;
			pnt2.y = -surface[x2][y2].y;
			Point3D_To_ScreenCoords(&pnt2, &loc2, x2, y2); // следующая точка

			x3 = x; y3 = z+1; ChangeIndexies_Painter(x3, y3);
			pnt3.x = surface[x3][y3].x;
			pnt3.z = surface[x3][y3].z;
			pnt3.y = -surface[x3][y3].y;
			Point3D_To_ScreenCoords(&pnt3, &loc3, x3, y3); // точка выше

			x4 = x+1; y4 = z+1; ChangeIndexies_Painter(x4, y4);
			pnt4.x = surface[x4][y4].x;
			pnt4.z = surface[x4][y4].z;
			pnt4.y = -surface[x4][y4].y;
			Point3D_To_ScreenCoords(&pnt4, &loc4, x4, y4); // точка выше и дальше

			DWORD color;

//			color = TrColor(&pnt1, &pnt2, &pnt3);
			color = TrColor(&loc1, &loc2, &loc3);

			Triangle(&IPoint(x1, y1), &IPoint(x2, y2), &IPoint(x3, y3), color);
			Triangle(&IPoint(x3, y3), &IPoint(x2, y2), &IPoint(x4, y4), color);

			Canvas->Pen->Color = clWhite;
			Canvas->MoveTo(x3, y3);
			Canvas->LineTo(x1, y1);
			Canvas->LineTo(x2, y2);
			Canvas->LineTo(x4, y4);
			Canvas->LineTo(x3, y3);
		}
	}
}

//######################################################################################################################
void CGraph3D::Translate(float x, float y, float z)
{
	cam_pos.x += x;
	cam_pos.y += y;
	cam_pos.z += z;
}

//######################################################################################################################
void CGraph3D::Rotate(float x, float y, float z)
{
	// отсечения вариантов области {(-0.5..0.5), y, (-0.5..0.5)}
	float new_x = cam_ang.x + x;
	if ((new_x >= -0.5) && (new_x <= 0.5)) cam_ang.x += x;
	
	cam_ang.y += y;
	cam_ang.z += z; 
}
