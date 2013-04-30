#include "CubeObj.h"

CCube::CCube()
{
	data = new TRIPLE[8];

	data[0] = Triple(-1, -1, -1);
	data[1] = Triple( 1, -1, -1);
	data[2] = Triple( 1, -1,  1);
	data[3] = Triple(-1, -1,  1);

	data[4] = Triple(-1,  1, -1);
	data[5] = Triple( 1,  1, -1);
	data[6] = Triple( 1,  1,  1);
	data[7] = Triple(-1,  1,  1);

	color = 0xFFFFFF;
}

void CCube::SetColor(const COLOR c)
{
	color = c & 0xFFFFFF;
}

void CCube::Draw()
{
	// !!! нормали проверены !!!
/*
	const byte ind[12][3] = {	{0,2,1},{2,0,3}, // низ
								{4,5,6},{4,6,7}, // верх
								{7,3,0},{7,0,4}, // лево
								{1,6,5},{6,1,2}, // право
								{4,0,1},{5,4,1}, // перед
								{7,2,3},{2,7,6}  // зад
							};
*/

	const byte ind[12][3] = {	{0,1,2},{0,2,3}, // низ
								{4,6,5},{6,4,7}, // верх
								{7,0,3},{0,7,4}, // лево
								{1,5,6},{1,6,2}, // право
								{4,1,0},{4,5,1}, // перед
								{7,3,2},{7,2,6}  // зад
							};

	TRIPLE tr[3], tmp[3];

	COLOR tmp_color;

	for (byte i=0; i<12; i++)
	{
		for (byte j=0; j<3; j++)
		{
			f3d_Localize( &data[ind[i][j]], ViewMatrix, &tmp[j]);
			f3d_Project(&tmp[j], &tr[j]);
		}

		if (flags & F3DOF_USE_LIGHT)
		{
			tmp_color = color;
			if (isPlaneVisible(tmp, &tmp_color))
				f3d_Triangle3D(tr[0], tr[1], tr[2], tmp_color);
		}
		else
			if (isPlaneVisible(tmp, NULL))
				f3d_Triangle3D(tr[0], tr[1], tr[2], color);		
	}
}
