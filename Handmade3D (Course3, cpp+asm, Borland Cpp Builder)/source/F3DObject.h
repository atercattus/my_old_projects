/*******************************************************************
*                           Fast3DLibrary
*   Графическая библиотека для работы с 2D/3D графикой.
*   Не используются OpenGL, DirectX, Glide, SDL, ets.
*   Основана на Win32API. Для Borland C++ Builder 6.
*
*                           F3DObject.*
*   Реализация базового класса 3D-объекта, поддерживающего
*   основные операции и свойства. Является мостом между
*   своими потомками и функциями библиотеки.
*******************************************************************/
#ifndef F3DObjectH
#define F3DObjectH

#include <windows.h>
#include "F3DL.h"
#include "F3DProj.h"

// ######## флаги свойств объекта ########

	// использовать плоское освещение по Ламберту
#define		F3DOF_USE_LIGHT				1
	// применять цвет источника освешения ("цветное" освещение)  
#define		F3DOF_USE_COLOR_LIGHT		2
	// применять направление источника света  
#define		F3DOF_USE_LIGHT_DIR			4


	class CF3DObject
	{
		protected:
			// видовая матрица
			double	ViewMatrix[16];

			// массив вершин и их число
			TRIPLE	*data;
			U32		data_len;

            // флаги (параметры вывода и т.д.)
			U32		flags;

			// параметры источника света (точка расположена в бесконечности)
			// используются, если выставлены флаги F3DOF_USE_LIGHT и F3DOF_USE_COLOR_LIGHT
			COLOR	light_color;		// цвет, по-умолчанию белый (0xFF, 0xFF, 0xFF)
			double	light_luminance;	// яркость, по-умолчанию 0.2
			TRIPLE	light_direction;	// направление, по-умолчанию вперед (0, 0, 1)

			// получение цвета плоскости pl с учетом освещенности,
			// если она имеет цвет материала mat
			COLOR getPlaneColor(const TRIPLE pl[3], COLOR &mat);

			// Проверка на видимость плоскости (учет нормали)
			// Если плоскость видна, то можно получить ее цвет
			//  с учетом освешенности. Для этого color!=NULL
			//  и должен быть равен абсолютному цвету плоскости.
			//  Тогда, в color сохраняется освещенность плоскости.  
			bool isPlaneVisible(const TRIPLE pl[3], COLOR *color);

		public:
			CF3DObject();
			~CF3DObject();

			// реализация рисования зависит от формата data
			virtual void Draw()=0;

            // восстановление начального состояния
			virtual void Restore();

			// пространственные операции
			virtual void Translate(double dx, double dy, double dz);
			virtual void Scale(double dx, double dy, double dz);
			virtual void Rotate(double dx, double dy, double dz);

            // работа с флагами
			virtual inline U32 GetFlags();
			virtual void SetFlags(U32 f);

            // работа с источником света
			U32 SetLightColor(COLOR &c);
			COLOR GetLightColor();

			U32 SetLumLight(double &lum);
			double GetLumLight();

			U32 SetDirLight(TRIPLE &d);
			TRIPLE GetDirLight();
	};


#endif // F3DObjectH
