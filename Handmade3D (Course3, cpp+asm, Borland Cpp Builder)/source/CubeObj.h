#ifndef CubeObjH
#define CubeObjH

#include "F3DObject.h"

	class CCube : public CF3DObject
	{
		protected:
			COLOR color;
		public:
			CCube();
			virtual void Draw();
			void SetColor(const COLOR c);
	};

#endif // CudeObjH
