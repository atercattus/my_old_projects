#include "F3DObject.h"

//######################################################################################################################
CF3DObject::CF3DObject()
{
	data = NULL;
	data_len = 0;

	flags = 0; // может флаги перенести в Restore() ?

	Restore();
}

//######################################################################################################################
CF3DObject::~CF3DObject()
{
	if (data_len) delete[] data; // а не "if (data)" !!!
}

//######################################################################################################################
void CF3DObject::Restore()
{
	f3d_MakeSingle(ViewMatrix);

	light_color = 0xFFFFFF;
	light_luminance = 0.2;
	light_direction = Triple(0.0, 0.0, 1.0);
}

//######################################################################################################################
void CF3DObject::Translate(double dx, double dy, double dz)
{
	f3d_Translate(ViewMatrix, ViewMatrix, Triple(dx, dy, dz));
}

//######################################################################################################################
void CF3DObject::Scale(double dx, double dy, double dz)
{
	double tmp[16];
	f3d_Scale(ViewMatrix, ViewMatrix, Triple(dx, dy, dz));
}

//######################################################################################################################
void CF3DObject::Rotate(double dx, double dy, double dz)
{
	f3d_Rotate(ViewMatrix, ViewMatrix, Triple(dx, dy, dz));
}

//######################################################################################################################
U32 CF3DObject::GetFlags()
{
	return flags;
}

//######################################################################################################################
void CF3DObject::SetFlags(U32 f)
{
	flags = f; 
}

//######################################################################################################################
bool CF3DObject::isPlaneVisible(const TRIPLE pl[3], COLOR *color)
{
//	const byte light_color[3] = {255, 0, 0};
//	const double light_color_per = 0.2;

    // ищу два вектора в плоскости
	TRIPLE a, b, n, f;
	a.x = pl[1].x-pl[0].x;
	a.y = pl[1].y-pl[0].y;
	a.z = pl[1].z-pl[0].z;

	b.x = pl[2].x-pl[0].x;
	b.y = pl[2].y-pl[0].y;
	b.z = pl[2].z-pl[0].z;

    // поиск нормали к плоскости
	f3d_Normal(&a, &b, true, &n);

    // смотрю вперед
	f = Triple(0, 0, 1);

	// угол между направлением взгляда и нормалью к плоскости
	double angle;
	f3d_Angle(&n, &f, &angle);

	// видимость плоскости
	bool res = ((angle > 92 ) && (angle < 268)); // реально нужно (90..270)

	if (!(flags & F3DOF_USE_LIGHT)) return res; // если свет не включен, то неча его обрабатывать

	if (res && color) // плоскость видна и нужно считать освещенность
	{
		*color &= 0xFFFFFF;

		// пересчет угла с учетом направления источника света
		if (flags & F3DOF_USE_LIGHT_DIR)
			f3d_Angle(&n, &light_direction, &angle);

		if (angle < 90 || angle > 270) angle = 0.03; // почти не видно (можно и константу ввести)
		else
		{
			angle -= 90.0; angle /= 90.0;
		}

		byte _r, _g, _b;

		// выделяю цветовые составляющие
		_r = *color >> 16,
		_g = (*color >> 8) & 0xFF,
		_b = *color & 0xFF;

		// пересчет с учетом параметров источника света
		if (flags & F3DOF_USE_COLOR_LIGHT)
		{
        	// получаю теперь угол к источнику света, а не к направлению взгляда

			double one_minus_lum = 1.0 - light_luminance;

			_r = ((double)_r*one_minus_lum + (double)(light_color >> 16)*light_luminance);
			_g = ((double)_g*one_minus_lum + (double)((light_color >> 8) & 0xFF)*light_luminance);
			_b = ((double)_b*one_minus_lum + (double)(light_color & 0xFF)*light_luminance);
		}

		// пересчет цветовых составляющих
		_r *= angle;
		_g *= angle;
		_b *= angle;

		// формирование нового цвета
		*color = (_r << 16) + (_g << 8) + _b;
	}

	return res;
}

//######################################################################################################################
U32 CF3DObject::SetLightColor(COLOR &c)
{
	light_color = (c & 0xFFFFFF); 
}

//######################################################################################################################
COLOR CF3DObject::GetLightColor()
{
	return light_color;
}

//######################################################################################################################
U32 CF3DObject::SetLumLight(double &lum)
{
	if (lum < 0.0 | lum > 1.0) return F3D_DATA_ERROR;

	light_luminance = lum;

    return F3D_NO_ERROR;
}

//######################################################################################################################
double CF3DObject::GetLumLight()
{
	return light_luminance;
}

//######################################################################################################################
U32 CF3DObject::SetDirLight(TRIPLE &d)
{
	double len = sqrt(d.x*d.x + d.y*d.y + d.z*d.z);
	if (!len) return F3D_DATA_ERROR;

	light_direction = d;

	return F3D_NO_ERROR;
}

//######################################################################################################################
TRIPLE CF3DObject::GetDirLight()
{
	return light_direction;
}
