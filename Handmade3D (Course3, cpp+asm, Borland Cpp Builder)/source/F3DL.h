/*******************************************************************
*                           Fast3DLibrary
*   ����������� ���������� ��� ������ � 2D/3D ��������.
*   �� ������������ OpenGL, DirectX, Glide, SDL, ets.
*   �������� �� Win32API. ��� Borland C++ Builder 6.
*
*                              F3DL.*
*   ������� ������ ����������, �������� ������� ����� ��������,
*   ����� � �������. ��������� ��� ������ �������.
*******************************************************************/
#ifndef F3DL_H
#define F3DL_H

#include <windows.h>
#include <fastmath.h>
#include <values.h>

//#################################### ������������ ���� ####################################################

    // �����, � ��� ������� { ������ �� ��������������, ������������ ��-��������� }
//	#define DOUBLE_BUFFERED

	// ��� ������ ����� ������������ �������� �����������, ����� �������� "J.Boyer & J.J.Bourdin" ( 8] ������...)
	#define BREZENHEM_LINE

	// �������� �������� ����������
	#define BOUNDS_AUTO_CHECK

    // �������� �� ����������� ������ ����� ������� ������������� ������� (e.g. ����� �� ������ �������� - ��� �����)
	#define CALL_OPTIMIZE

	// ���������� �������� ��� ����������� � ��������� ������� { ������ �� �������������� }
//	#define NO_REC_BOUNDS_CHECK


//###########################################################################################################
//#################################### ��������� ���������� #################################################

#define	F3D_NO_ERROR			 0	// ��� ������
#define	F3D_DATA_ERROR			-1	// ������������ ���������
#define	F3D_CRAZY_CALL			-2	// ������� ������������ � ������ ������ ������� (�� ������, � �����������)
#define	F3D_API_ERROR			-3	// ������ ��� ������ API-������� (������ ����� ������ ������)
#define	F3D_MEMORY_ERROR		-4	// ������������ ������ ��� ��������� ������ ��� ������ � �������



//###########################################################################################################
//#################################### ��������� ############################################################

	// �������� ������������� ������ �����
#define	F3D_MAX_COLOR			0xFFFFFF



//###########################################################################################################
//#################################### ��������������� ������� ##############################################

#define f3d_RGB(r,g,b) ((((r)<<8) | (g)) << 8) | (b)
#define f3d_min(a,b) ((a)<(b)) ? (a) : (b)
#define f3d_max(a,b) ((a)>(b)) ? (a) : (b)
#define f3d_abs(a) ((a)<0) ? -(a) : (a)

#define S32 signed long int
#define U32 unsigned long int
#define COLOR U32

struct POINT32
{
	S32 x, y;
};

struct TRIPLE
{
	double x, y, z;
};

//###########################################################################################################
//################################## ����������� ���������� ( 2D )###########################################

		// ������������� �������
	U32 f3d_Init(HDC _hDC, U32 _width, U32 _height);

		// ������������ ��������
	U32 f3d_Done();

		// ������������ �������
	U32 f3d_Flip();

		// ������� ������ �����
	U32 f3d_Clear(COLOR color);

		// ������� ������ �����. �������, ��� f3d_Clear, �� ���� ����� ������ (�.�. ������� ������)
	U32 f3d_ClearByte(BYTE color);

		// ����� �������
	U32 f3d_SetPixel(U32 x, U32 y, U32 color);

    	// ���������� ������� (���������� ���� ����, ���� ������)
	U32 f3d_GetPixel(U32 x, U32 y);

    	// ����� ����� � ��������� ������
	U32 f3d_Line2D(U32 x1, U32 y1, U32 x2, U32 y2, COLOR color);

    	// ����� �������������� �����
	U32 f3d_HLine(S32 x1, S32 x2, S32 y, COLOR color);

		// ����� ������������ �����
	U32 f3d_VLine(S32 x, S32 y1, S32 y2, COLOR color);

    	// ����� ������������ � ���������
	U32 f3d_Triangle(POINT32 p1, POINT32 p2, POINT32 p3, COLOR color);

//###########################################################################################################
//################################## ����������� ���������� ( 3D )###########################################

		// ������������� ������ �������
	U32 f3d_ZbufInit();

		// ������������ �������� ��� ����� �������
	U32 f3d_ZbufDone();

		// ������� ������ �������
	U32 f3d_ZbufClear();

		// ������������ ���������
	TRIPLE Triple(double x, double y, double z);

		// ����� ����� ������������ ����� � ������ ������ �������
	U32 f3d_SetPixel3D(S32 x, S32 y, double z, COLOR color);

		// ����� �������������� (� �������� �� �����) �����
	U32 f3d_HLine3D(S32 x1, S32 x2, double z1, double z2, S32 y, COLOR color);

		// ����� ������������ � ������������ (������� ���������, ���� �������������� �� �������)
	U32 f3d_Triangle3D(TRIPLE &p1, TRIPLE &p2, TRIPLE &p3, COLOR color);

//###########################################################################################################
//#################################### �������������� ������� ###############################################

		// ��������������� ���� ��� �������� ������� ���������� �������. ����� ��� � ���������� �� ������.
	void f3d_SetClientRect(HWND hWnd, U32 width, U32 height, bool centered);

		// ������������ �������� �� 4.
	U32 f3d_Align4(U32 val);

    	// ������� dest ������ len 24-������ ������ 24-������ ��������� color
	void f3d_memset_24(U32 dest, U32 len, COLOR color);

		// ���������� ����� ��� unsigned long int
	S32 f3d_sign(S32 a);

		// ���������� � ������������� ������� �����
	double f3d_ftrunc(double a);

    	// ���������� �� ������, ������������� ������� ����� 
	S32 f3d_fround(double a);

    	// ������������� ��������� ������ [s] � �������� [c] ���� [a], ��������� � ��������
	void _sincos_(double a, double &s, double &c);

    	// ���������� ����������� a ���� c
	void _arccos_(double *c, double *a);

    	// ������������ ��������� POINT32 �� ��������� ����������
	POINT32 Point32(S32 x, S32 y);

    	// ����� ��������
	void swap(TRIPLE &v1, TRIPLE &v2);

//###########################################################################################################
//################################## ������ �� ���������� ����� #############################################

    	// ��������� ������� ������� ������
	void f3d_GetBuffDims(POINT32 *res);
	void f3d_GetBuffHalfDims(POINT32 *res);

#pragma warn -rvl // �������� �� ���������� ������������� ��������

#endif // F3DL_H
