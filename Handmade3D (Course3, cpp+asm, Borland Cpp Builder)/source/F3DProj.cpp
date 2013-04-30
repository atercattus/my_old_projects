#include "F3DProj.h"

/******************************** Local constants **********************************/

	const DOUBLE_16 = (MAT_SIZE*sizeof(double));
	const DOUBLE_3 = (3*sizeof(double));

	const double PI180 = 0.017453292519943295769236907684886; /* M_PI/180.0 */
	const double VAL180 = 180.0;

	
//######################################################################################################################
U32 f3d_MakeSingle(double m[MAT_SIZE])
{
	const double single[16] = {	1, 0, 0, 0,
								0, 1, 0, 0,
								0, 0, 1, 0,
								0, 0, 0, 1
								};

	CopyMemory((void*)m, (void*)single, DOUBLE_16);

	return F3D_NO_ERROR;
}

//######################################################################################################################
U32 f3d_CopyMatrix(const double s[MAT_SIZE], double d[MAT_SIZE])
{
	RtlCopyMemory((void*)d, (void*)s, DOUBLE_16);

	return F3D_NO_ERROR;
}

//######################################################################################################################
U32 f3d_MulMatrix(const double A[MAT_SIZE], const double B[MAT_SIZE], double AB[MAT_SIZE])
{
	double *rp = AB;

	for (int i=0; i<4; i++)
	{
		for (int j=0; j<4; j++)
		{
			*rp = 0;
			double *Ap = (double*)&A[i << 2];
			double *Bp = (double*)&B[j];

			for (int k=0; k<4; k++)
			{
				*rp += (*Ap++)*(*Bp);
				Bp += 4;
			}
			rp++;
		}
	}

    return F3D_NO_ERROR;
}

//######################################################################################################################
U32 f3d_Project(const TRIPLE *P, TRIPLE *R)
{
	double zd = P->z;

	POINT32 res;
	f3d_GetBuffHalfDims(&res);

	if (zd != -camera_focus)	zd = camera_focus / (zd + camera_focus);
	else						zd = 1.0;

	R->x = res.x * (1 + P->x * zd);
	R->y = res.y * (1 + P->y * zd);
	R->z = P->z;

	return F3D_NO_ERROR;
}

//######################################################################################################################
U32 f3d_Localize(const TRIPLE *P, const double VM[MAT_SIZE], TRIPLE *R)
{
/****************************************
   VARIANT 0
*

	double vi;

	for (int i=0; i<4; i++)
	{
		vi = 0;
		for (int j=0; j<4; j++)
			vi += ( VM[(i << 2)+j] * pnt[j] );
		tmp[i] = vi;
	}

*
***************************************/

/****************************************
   VARIANT 1
*

	double vi;

	for (int i=0; i<4; i++)
	{
		vi = 0;
		for (int j=0; j<3; j++)
			vi += ( VM[(i << 2)+j] * pnt[j] );
		tmp[i] = (vi + VM[(i << 2)+3]);
	}

*
***************************************/

/****************************************
   VARIANT 2
*
   
	for (int i=0; i<4; i++)
	{
		tmp[i] =	( VM[(i << 2)]   * pnt[0] ) +
					( VM[(i << 2)+1] * pnt[1] ) +
					( VM[(i << 2)+2] * pnt[2] ) +
					( VM[(i << 2)+3] );
	}

*
***************************************/

/****************************************
   VARIANT 3
*

	tmp[0] = (VM[0] *  pnt[0]) + (VM[1] *  pnt[1]) + (VM[2] *  pnt[2]) + VM[3];
	tmp[1] = (VM[4] *  pnt[0]) + (VM[5] *  pnt[1]) + (VM[6] *  pnt[2]) + VM[7];
	tmp[2] = (VM[8] *  pnt[0]) + (VM[9] *  pnt[1]) + (VM[10] * pnt[2]) + VM[11];
	tmp[3] = (VM[12] * pnt[0]) + (VM[13] * pnt[1]) + (VM[14] * pnt[2]) + VM[15];

*	
***************************************/

/****************************************
   VARIANT 4
*

	tmp[0] = (VM[0] *  (*pntr)) + (VM[1] *  (*(pntr+1))) + (VM[2] *  (*(pntr+2))) + VM[3];
	tmp[1] = (VM[4] *  (*pntr)) + (VM[5] *  (*(pntr+1))) + (VM[6] *  (*(pntr+2))) + VM[7];
	tmp[2] = (VM[8] *  (*pntr)) + (VM[9] *  (*(pntr+1))) + (VM[10] * (*(pntr+2))) + VM[11];
	tmp[3] = (VM[12] * (*pntr)) + (VM[13] * (*(pntr+1))) + (VM[14] * (*(pntr+2))) + VM[15];

*
***************************************/

/****************************************
   VARIANT 5
*/

	double *tmp = (double*)R;

	// tmp = A + B + C + D

	asm
	{
		mov		esi,	P
		mov		edi,	tmp
		mov		edx,	VM

// вычисл€ю все четыре произведени€ A

		fldz
		fldz
		fldz
		fldz

		add		edx,	11*8
        sub		esi,	8

// in stack: 0, 0, 0, 0

		xor		ecx,	ecx
		inc		ecx
		inc		ecx			// ecx = 2

@@loop:

// ¬ычисл€ю все четыре произведени€ A, B, C
//  омментарии в цикле даны дл€ второй итерации

		// edx корректирую на VM[1,2]
		sub		edx,	11*8

		// esi корректирую на P[1,2]
		add		esi,	8

		// VM[1] * P[1] { !!! B0 !!! }
		fld		qword ptr [esi]
		fld		st(0)
		fld		qword ptr [edx]
		fmulp
		faddp	st(5), st(0) // { !!! A0 + B0 !!!}

		// VM[5] * P[1] { !!! B1 !!! }
		fld		st(0)
		add		edx,	4*8
		fld		qword ptr [edx]
		fmulp
		faddp	st(4), st(0) // { !!! A1 + B1 !!!}
		
		// VM[9] * P[1] { !!! B2 !!! }
		fld		st(0)
		add		edx,	4*8
		fld		qword ptr [edx]
		fmulp
		faddp	st(3), st(0) // { !!! A2 + B2 !!!}

		// VM[13] * P[1] { !!! B3 !!! }
		add		edx,	4*8
		fld		qword ptr [edx]
		fmulp
		// внимание!!!! тут именно "st(1)", а не "st(2)", т.к. € не копировал "fld st(0)" !!!!!
		faddp	st(1), st(0) // { !!! A3 + B3 !!!}

// in stack: D1+D0, C1+C0, B1+B0, A1+A0

		test	ecx,	ecx
		jz		@@exit

		dec		ecx
		jmp		@@loop

@@exit:

// in stack: D2+D1+D0, C2+C1+C0, B2+B1+B0, A2+A1+A0

// осталось добавить x3

		// edx correct on VM[3]
		sub		edx,	11*8

		// D0 + (C0+B0+A0)
		fld		qword ptr [edx]
		faddp	st(4), st(0)

		// ...
		fld		qword ptr [edx+4*8]
		faddp	st(3), st(0)

		// ...
		fld		qword ptr [edx+2*4*8]
		faddp	st(2), st(0)

        // ...
		fld		qword ptr [edx+3*4*8]
		faddp	st(1), st(0)

// теперь сохран€ю результаты в R
		fstp	qword ptr [edi]		// нафига € считаю x3, если не использую. да и почему не использую то???!!!
		fstp	qword ptr [edi+2*8]
		fstp	qword ptr [edi+1*8]
		fstp	qword ptr [edi]
	}

/*
***************************************/

	return F3D_NO_ERROR;
}

//######################################################################################################################
U32 f3d_Translate(const double VM[MAT_SIZE], double VMNew[MAT_SIZE], TRIPLE d)
{
	const
	double matrix[MAT_SIZE] = {		1,		0,		0,		d.x,
									0,		1,		0,		d.y,
									0,		0,		1,		d.z,
									0,		0,		0,		1
								};

	double tmp[MAT_SIZE];

	f3d_MulMatrix(VM, matrix, tmp);

	f3d_CopyMatrix(tmp, VMNew);

	return F3D_NO_ERROR;
}

//######################################################################################################################
U32 f3d_Scale(const double VM[MAT_SIZE], double VMNew[MAT_SIZE], TRIPLE d)
{
	const
	double matrix[MAT_SIZE] = {		d.x,	0,		0,		0,
									0,		d.y,	0,		0,
									0,		0,		d.z,	0,
									0,		0,		0,		1
								};

	double tmp[MAT_SIZE];

	f3d_MulMatrix(VM, matrix, tmp);

	f3d_CopyMatrix(tmp, VMNew);

	return F3D_NO_ERROR;
}

//######################################################################################################################
U32 f3d_Rotate(const double VM[MAT_SIZE], double VMNew[MAT_SIZE], TRIPLE d)
{
/*
	VMNew = RotateX * RotateY * RotateZ * VM;
	"Rotate{A}" == "d.{a}"
	Rotate{A} in degrees
*/

	double s_a, c_a, s_b, c_b, s_g, c_g;

	_sincos_(d.x, s_a, c_a);
	_sincos_(d.y, s_b, c_b);
	_sincos_(d.z, s_g, c_g);

	const double matrix[MAT_SIZE] =	{

							c_b * c_g,							c_b * s_g,				  -s_b,			0,
				(s_a * s_b * c_g) - (c_a * s_g),	(s_a * s_b * s_g) + (c_a * c_g),	s_a * c_b,		0,
				(c_a * s_b * c_g) + (s_a * s_g),	(c_a * s_b * s_g) - (s_a * c_g),	c_a * c_b,		0,
								0,									0,						0,			1

									};


	double tmp[MAT_SIZE];

	f3d_MulMatrix(VM, matrix, tmp);

	f3d_CopyMatrix(tmp, VMNew);

	return F3D_NO_ERROR;
}

//######################################################################################################################
U32 f3d_Normal(const TRIPLE *a, const TRIPLE *b, bool normalize, TRIPLE *n)
{
	n->x = (a->y * b->z) - (a->z * b->y);
	n->y = (a->z * b->x) - (a->x * b->z);
	n->z = (a->x * b->y) - (a->y * b->x);

	if (normalize)
	{
		asm
		{
			mov		eax,	n

			fld		qword ptr [eax]
			fld		st(0)
			fmulp

			fld		qword ptr [eax+8]
			fld		st(0)
			fmulp

			fld		qword ptr [eax+16]
			fld		st(0)
			fmulp

			faddp
			faddp
			fsqrt

			fld1

			fucomp	st(1)
			fstsw	ax
			sahf
			jz @@exit

				mov		eax,	n

				fld		st(0)
				fld		qword ptr [eax]
				fdivr
				fstp	qword ptr [eax]

				add		eax,	8
				fld		st(0)
				fld		qword ptr [eax]
				fdivr
				fstp	qword ptr [eax]

				add		eax,	8
				fld		qword ptr [eax]
				fdivr
				fstp	qword ptr [eax]
		@@exit:
				fninit	// удал€ю длину нормали из st(0)
		}
	}

    return F3D_NO_ERROR;
}

//######################################################################################################################
U32 f3d_Angle(const TRIPLE *a, const TRIPLE *b, double *angle)
{
	if (!angle) return F3D_CRAZY_CALL;

	double sqr_len_a = (a->x * a->x) + (a->y * a->y) + (a->z * a->z);
	if (!sqr_len_a) return F3D_DATA_ERROR;

	double sqr_len_b = sqrt((b->x * b->x) + (b->y * b->y) + (b->z * b->z));
	if (!sqr_len_b) return F3D_DATA_ERROR;
	
	double a_mul_b = (a->x * b->x) + (a->y * b->y) + (a->z * b->z);

	double an = a_mul_b / (sqr_len_a*sqr_len_b);

	_arccos_(&an, angle);

	*angle /= PI180; // все углы в программе в √–јƒ”—ј’ !!!
	
	return F3D_NO_ERROR;
}
