#include "F3DL.h"

/******************************** Local variables **********************************/

		// приемный контекст
	HDC			hDestDC;

		// собственный контекст
	HDC			hDC;

    	// собственный bitmap
	HBITMAP		hBitmap;

		// параметры bitmap'а
	BITMAPINFO	bitmap_info;

		// собственный буфер цвета
	BYTE		*lpColor;

    	// размер буфера цвета в пикселях
	U32		ColorBuffSize;

		// размер буфера цвета в байтах
	U32		ColorBuffSizeByte;
	
		// размер строки буфера цвета в байтах
	U32		ColorBuffLine;

		// размер области вывода в пикселях
	S32		BuffWidth,	BuffHeight;

		// размер половины области вывода в пикселях
	S32		BuffWidth2,	BuffHeight2;

		// буфер глубины (z точек УМНОЖЕННЫЕ НА Z_BUF_SCALE)
	double	*Zbuf = NULL;

		// флаг того, что z-координаты уже отмасштабированы
	bool	Z_BUF_SCALE_USED = false;

		// множитель глубин точек
	const double Z_BUF_SCALE = 1e100;

//######################################################################################################################
U32 f3d_Init(HDC _hDC, U32 _width, U32 _height)
{
	if (hDC)
	{ // система уже инициализирована
		return F3D_CRAZY_CALL;
	}

	if (!(GetDeviceCaps(_hDC, RASTERCAPS) & RC_BITBLT))
	{
		// в качестве HDC передано что-то другое, либо система не поддерживает BitBlt
		return F3D_DATA_ERROR;
	}

	// создаю свой совместимый
	hDC = CreateCompatibleDC(_hDC);
	if (!hDC) return F3D_API_ERROR;

	// сохраняю контекст
	hDestDC = _hDC;

	// формирую BITMAPINFO

	BuffWidth = f3d_Align4(_width);
	BuffHeight = _height;
	ColorBuffLine = BuffWidth * 3;

	BuffWidth2 = BuffWidth >> 1;
	BuffHeight2 = BuffHeight >> 1;

	memset(&bitmap_info, 0, sizeof(BITMAPINFO));

	bitmap_info.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
	bitmap_info.bmiHeader.biWidth = BuffWidth;
	bitmap_info.bmiHeader.biHeight = BuffHeight;
	bitmap_info.bmiHeader.biPlanes = 1;
//#ifdef 
	bitmap_info.bmiHeader.biBitCount = 24;
	bitmap_info.bmiHeader.biCompression = BI_RGB;

	// создаю буфер цвета
	hBitmap = CreateDIBSection(hDC, &bitmap_info, DIB_RGB_COLORS, (void**)&lpColor, NULL, 0);
	if (!lpColor || !hBitmap)
	{
    	return F3D_API_ERROR;
	}

	SelectObject(hDC, hBitmap);

	ColorBuffSize = BuffWidth * BuffHeight;
	ColorBuffSizeByte = ColorBuffSize*3;

	return F3D_NO_ERROR;
}

//######################################################################################################################
U32 f3d_Done()
{
	if (!hDC)
	{ // система не инициализирована
		return F3D_CRAZY_CALL;
	}

	DeleteObject(hBitmap);

	DeleteDC(hDC);

	hDC = NULL;

	return F3D_NO_ERROR;
}

//######################################################################################################################
U32 f3d_Flip()
{
	if (!hDC)
	{ // система не инициализирована
		return F3D_CRAZY_CALL;
	}

	BitBlt(hDestDC, 0, 0, BuffWidth, BuffHeight, hDC, 0, 0, SRCCOPY);

	return F3D_NO_ERROR;
}

//######################################################################################################################
U32 f3d_Clear(COLOR color)
{
	if (!hDC)
	{ // система не инициализирована
		return F3D_CRAZY_CALL;
	}

	f3d_memset_24((U32)lpColor, ColorBuffSize, color);

	return F3D_NO_ERROR;
}

//######################################################################################################################
U32 f3d_ClearByte(BYTE color)
{
	if (!hDC)
	{ // система не инициализирована
		return F3D_CRAZY_CALL;
	}

	memset(lpColor, color, ColorBuffSizeByte);

	return F3D_NO_ERROR;
}

//######################################################################################################################
U32 f3d_SetPixel(S32 x, S32 y, COLOR color)
{
	if (!hDC)
	{ // система не инициализирована
		return F3D_CRAZY_CALL;
	}

	asm
	{
		mov		eax,	y
#ifdef BOUNDS_AUTO_CHECK
		test	eax,	eax
		jl		@@error
		cmp		eax,	BuffHeight
		jnl		@@error
#endif // BOUNDS_AUTO_CHECK

		mov		ecx,	x
#ifdef BOUNDS_AUTO_CHECK
		test	ecx,	ecx
		jl		@@error
		cmp		ecx,	BuffWidth
		jnl		@@error
#endif // BOUNDS_AUTO_CHECK

		xor		edx,	edx
		mul		ColorBuffLine

		mov		edx,	ecx
		lea		edx,	[edx+2*edx]		// edx = x*3

		add		eax,	edx
		add		eax,	lpColor			// eax = lpColor + (ColorBuffLine*y) + (x*3) // полный адрес точки

		mov		ecx,	color

		mov		word ptr [eax],	cx
		inc		eax
		inc		eax
		shr		ecx,	16
		mov		byte ptr [eax], cl

		xor		eax,	eax
		inc		eax		// return eax = F3D_NO_ERROR

		pop		ebp
		ret

#ifdef BOUNDS_AUTO_CHECK
	@@error:	// выход за пределы окна

		xor		eax,	eax
		dec		eax			// Result = F3D_DATA_ERROR

		pop		ebp
		ret
#endif // BOUNDS_AUTO_CHECK
	}
}

//######################################################################################################################
U32 f3d_GetPixel(U32 x, U32 y)
{
	if (!hDC)
	{ // система не инициализирована
		return F3D_CRAZY_CALL;
	}

	asm
	{
		mov		eax,	y
#ifdef BOUNDS_AUTO_CHECK
		test	eax,	eax
		jl		@@error
		cmp		eax,	BuffHeight
		jnl		@@error
#endif // BOUNDS_AUTO_CHECK

		mov		ecx,	x
#ifdef BOUNDS_AUTO_CHECK
		test	ecx,	ecx
		jl		@@error
		cmp		ecx,	BuffHeight
		jnl		@@error
#endif // BOUNDS_AUTO_CHECK

		xor		edx,	edx
		mul		ColorBuffLine

		mov		edx,	ecx
		lea		edx,	[edx+2*edx]		// edx = x*3

		add		eax,	edx
		add		eax,	lpColor			// eax = lpColor + (ColorBuffLine*y) + (x*3) // полный адрес точки

		mov		cx,		word ptr [eax]
		shl		ecx,	8
		mov		cl,		byte ptr [eax]

		xchg	eax,	ecx	// Result = color

		pop		ebp
		ret

#ifdef BOUNDS_AUTO_CHECK
	@@error:	// выход за пределы окна

		xor		eax,	eax
		dec		eax			// Result = F3D_DATA_ERROR

		pop		ebp
		ret
#endif // BOUNDS_AUTO_CHECK		
	}
}

#ifdef BREZENHEM_LINE

U32 f3d_Line2D(U32 x1, U32 y1, U32 x2, U32 y2, COLOR color)
{
	if (!hDC)
	{ // система не инициализирована
		return F3D_CRAZY_CALL;
	}

	S32		errX, errY,
			dX, dY,
			incX, incY,
			d;

	errX = 0;
	errY = 0;

	dX = x2 - x1 + 1;
	dY = y2 - y1 + 1;

	// проверки диапазонов ...

#ifdef CALL_OPTIMIZE
	if (!dX)
	{
		if (!dY)	return f3d_SetPixel(x1, y1, color);
		else		return f3d_VLine(x1, y1, y2, color);
	} else
	{
		if (!dY)	return f3d_HLine(x1, x2, y1, color);
	}
#endif // CALL_OPTIMIZE

	incX = f3d_sign(dX);
	incY = f3d_sign(dY);

	dX = f3d_abs(dX);
	dY = f3d_abs(dY);

	d = f3d_max(dX, dY);

	f3d_SetPixel(x1, y1, color);

	for (S32 i = 1; i < d; i++) 
	{
		errX += dX;

		if (errX > d)
		{
			errX -= d;
			x1 += incX;
		}

		errY += dY;

		if (errY > d)
		{
			errY -= d;
			y1 += incY;
		}

        f3d_SetPixel(x1, y1, color);
	}

	return F3D_NO_ERROR;
}

#else // BREZENHEM_LINE

U32 f3d_Line2D(U32 x1, U32 y1, U32 x2, U32 y2, COLOR color)
{
	if (!hDC)
	{ // система не инициализирована
		return F3D_CRAZY_CALL;
	}

	if (y1>y2)
	{
		S32 tmp = x1;
		x1 = x2;
		x2 = tmp;

		tmp = y1;
		y1 = y2;
		y2 = tmp;
	}

#ifdef BOUNDS_AUTO_CHECK
	if (y1 >= BuffHeight) return F3D_DATA_ERROR;
	if (y2 >= BuffHeight) return F3D_DATA_ERROR; // temporary
	if (x1 >= BuffWidth)  return F3D_DATA_ERROR; // ??
	if (x2 >= BuffWidth)  return F3D_DATA_ERROR; // ??
#endif // BOUNDS_AUTO_CHECK

	S32		dX = f3d_abs(x2-x1),
			dY = y2 - y1,
			min_l = dX / dY,
			err = 0,
			d_err = dX - min_l*dY,
			incX = (x2>x1) ? 1 : -1;

	while (y1 <= y2)
	{

		for (S32 i=0; i<min_l; i++)
		{
			f3d_SetPixel(x1, y1, color);
			x1 += incX;
		}

		err += d_err;

		if (err >= dY)
		{
			err -= dY;
			f3d_SetPixel(x1, y1, color);
			x1 += incX;
		}

        y1++;
	}

	return F3D_NO_ERROR;
}

#endif // BREZENHEM_LINE

//######################################################################################################################
U32 f3d_HLine(S32 x1, S32 x2, S32 y, COLOR color)
{
	if (!hDC)
	{ // система не инициализирована
		return F3D_CRAZY_CALL;
	}

	if (x1>x2)
	asm
	{
		mov		eax,	x1
		xchg	eax,	x2
		xchg	eax,	x1
	}

#ifdef CALL_OPTIMIZE
	else
	if (x1 == x2)
	{
		f3d_SetPixel(x1, y, color);
		return F3D_NO_ERROR;
	}
#endif // CALL_OPTIMIZE

#ifdef BOUNDS_AUTO_CHECK

	if (x1 >= BuffWidth) return F3D_DATA_ERROR;
	if (y >= BuffHeight) return F3D_DATA_ERROR;

    x1 = f3d_max(x1, 0);
	x2 = f3d_min(x2, BuffWidth-1);

#endif // BOUNDS_AUTO_CHECK

	void *ptr = lpColor + y*ColorBuffLine + x1*3;

	f3d_memset_24((U32)ptr, x2-x1+1, color);

	return F3D_NO_ERROR;
}

//######################################################################################################################
U32 f3d_VLine(S32 x, S32 y1, S32 y2, COLOR color)
{
	if (!hDC)
	{ // система не инициализирована
		return F3D_CRAZY_CALL;
	}

	asm
	{
		mov		eax,	y1
		mov		edx,	y2
		cmp		eax,	edx
		jna		@@ok
			// y1 > y2
			xchg	eax,	edx
			// я теряю значения !!!
	@@ok:

#ifdef BOUNDS_AUTO_CHECK

	// проверка диапазонов !!!

		cmp		eax,	BuffHeight

#endif // BOUNDS_AUTO_CHECK

	mov		y1,		eax
	mov		y2,     edx

	}

#ifdef CALL_OPTIMIZE

	if (y1 == y2)
		return f3d_SetPixel(x, y1, color);

#endif // CALL_OPTIMIZE

	asm
	{
		mov		eax,	y1

		push	ebx

		mov		ebx,	y2
		sub		ebx,	eax
		inc		ebx

		mov		ecx,	color

		xor		edx,	edx
		mul		ColorBuffLine

		mov		edx,	x
		lea		edx,	[edx+2*edx]		// edx = x*3

		add		eax,	edx
		add		eax,	lpColor			// eax = lpColor + (ColorBuffLine*y) + (x*3) // полный адрес точки

	@@loop:
		mov		word ptr [eax],	cx
		inc		eax
		inc		eax

		mov		edx,	ecx
		shr		edx,	16

		mov		byte ptr [eax], dl

		add		eax,	ColorBuffLine
		dec		eax
		dec		eax

		dec		ebx
		jnz		@@loop

		pop		ebx
	}

	return F3D_NO_ERROR;
}

//######################################################################################################################
void swap(POINT32 &p1, POINT32 &p2)
{
	POINT32 tmp = p1; p1 = p2; p2 = tmp;
}

//######################################################################################################################
void f3d_SetClientRect(HWND hWnd, U32 width, U32 height, bool centered)
{
	RECT Size, Rect;

	// клиентская область
	GetClientRect(hWnd, &Rect);
	++Rect.right;
	++Rect.bottom;

	// исходный размер окна
	GetWindowRect(hWnd, &Size);
	Size.right -= Size.left;
	Size.bottom -= Size.top;

    // новый размер окна
	Size.right += (width - Rect.right+1);
	Size.bottom += (height - Rect.bottom+1);

	// если нужно, центрирую
	if (centered)
	{
		Size.left = (GetSystemMetrics(SM_CXSCREEN) - Size.right) >> 1;
		Size.top = (GetSystemMetrics(SM_CYSCREEN) - Size.bottom) >> 1;
	}

	MoveWindow(hWnd, Size.left, Size.top, Size.right, Size.bottom, true);
}

//######################################################################################################################
U32 f3d_Align4(U32 val)
{
	asm
	{
		mov		eax,	val
		test	eax,	eax
		jne		@@no_zero
			mov		al,		4 // eax = 4
			jmp short @@exit
	@@no_zero:

		test	al,	3
		jz		@@exit
			and		al,		not 3
			add		eax,	4
	@@exit:
	}
}

//######################################################################################################################
void f3d_memset_24(U32 dest, U32 len, COLOR color)
{
	asm
	{
		push	edi

		mov		ecx,	len
		mov		edi,	dest
		test	ecx,	ecx
		je		@@exit
		dec		ecx
		jz		@@len_is_one
		mov		eax,	color

@@loop:
		mov		dword ptr [edi],	eax

		add		edi,	3

/*
 если описать цикл через loop @@loop, то у меня FPS=388
 если расписать так, то FPS=412(!)
 т.е. прирост на 24(!!!) кадра (600x600x24) при увеличении кода всего на 1 байт.
*/
		dec		ecx
		jnz		@@loop

@@len_is_one:

		// с последними 3 байтами приходится быть аккуратнее...
		mov		word ptr [edi],		ax
		shr		eax,	16
		mov		byte ptr [edi+2],	al

@@exit:

		pop		edi
	}
}

//######################################################################################################################
S32 f3d_sign(S32 a)
{
	asm
	{
		xor		eax,	eax // eax = 0 <- подготовка
		cmp		a,		0
		jg		@@g
		jl		@@l

		jmp		@@exit // sign = 0

	@@g:
		inc		eax
		jmp		@@exit

	@@l:
		dec		eax

	@@exit:
	}
}

//######################################################################################################################
void f3d_GetBuffDims(POINT32 *res)
{
	res->x = BuffWidth;
	res->y = BuffHeight;
}

//######################################################################################################################
void f3d_GetBuffHalfDims(POINT32 *res)
{
	res->x = BuffWidth2;
	res->y = BuffHeight2;
}

//######################################################################################################################
double f3d_ftrunc(double a)
{
	double r;
	short int mem16;
	asm
	{
		fld		a
		fstcw	mem16
		and		mem16,	1111001111111111b
		or		mem16,	1111011111111111b
		fldcw	mem16
		frndint       
		fstp	r	
	}
	return r;
}

//######################################################################################################################
S32 f3d_fround(double a)
{
	S32 r;
	short int mem16;
	asm
	{
		fld		a
		fstcw	mem16
		and		mem16,	1111001111111111b
		or		mem16,	1111011111111111b
		fldcw	mem16
		frndint
		fistp	r	
	}
	return r;
}

//######################################################################################################################
void _sincos_(double a, double &s, double &c)
{
	// уже объявил в F3DProj, но, что б не парится с сылками, вынес сюда 
	const double _PI_180_ = 0.017453292519943295769236907684886; /* M_PI/180.0 */
	asm
	{
		fld			a
		fld			_PI_180_
		fmul
		fsincos

		FWAIT // ?

		mov			eax,	c
		FSTP        qword ptr [eax]

		mov			eax,	s
		FSTP        qword ptr [eax]
	}
}

//######################################################################################################################
void _arccos_(double *c, double *a)
{
/*

Delphi6: "\Source\Rtl\Common\Math.pas"
-------------------------------------------------

	ArcCos(C) := ArcTan2(Sqrt(1 - C * C), C);

	function ArcTan2(const Y, X: Extended): Extended;
	asm
			FLD     Y
			FLD     X
			FPATAN
			FWAIT
	end;

-------------------------------------------------

*/

	asm
	{
		push	eax // ?

		mov		eax,	c
		fld		qword ptr [eax]
		fld		st(0)
		fmulp
		fld1
		fsubrp
		fsqrt
		fld		qword ptr [eax]
		fpatan
		fwait // ?
		mov		eax,	a
		fstp	qword ptr [eax]

		pop		eax // ?
	}
}

//######################################################################################################################
POINT32 Point32(S32 x, S32 y)
{
	POINT32 p;
	p.x = x;
	p.y = y;
	
	return p;
}

//######################################################################################################################
U32 f3d_Triangle(POINT32 p1, POINT32 p2, POINT32 p3, COLOR color)
{
	if (!hDC)
	{ // система не инициализирована
		return F3D_CRAZY_CALL;
	}

    // сортировка вершин
	if (p1.y > p2.y)
		// поменять местами вершины 1 и 2
		swap(p1, p2);
	if (p2.y > p3.y)
		// поменять местами вершины 2 и 3
		swap(p2, p3);
	if (p1.y > p2.y)
		// поменять местами вершины 1 и 2
		swap(p1, p2);

	S32 x1, x2, x3, sy, p1x, p2x, sy_s, sy_f, sy_py, sy_py2;

	double k12, k13, k23;

	if ((p3.y - p1.y) == 0) return F3D_DATA_ERROR; // вырожденный треугольник
	k13 = (double)(signed(p3.x) - signed(p1.x)) / (double)(signed(p3.y) - signed(p1.y));

    // указываю, что необходимо округлять числа в меньшую сторону
	short int mem16;
	asm
	{
		fstcw	mem16
		and		mem16,	1111001111111111b
		or		mem16,	1111011111111111b
		fldcw	mem16
	}

	p1x = p1.x;
	p2x = p2.x;

	if (!(p2.y - p1.y) || (p2.y < 0)) goto out_2_part; // нет нижней половины, либо она не видна
	k12 = (double)(p2.x - p1.x) / (double)(p2.y - p1.y);

// Нижняя часть

	sy_py = 0;

	sy_s = f3d_max(p1.y, 0);
	sy_f = f3d_min(p2.y, BuffHeight-1);

	for (sy = sy_s; sy <= sy_f; sy++, sy_py++)
	{
		asm
		{
			mov		ecx,	p1x

//		x1 = p1.x + f3d_fround(sy_py * k13);
			fild	sy_py
			fld		st(0)
			fld		k13
			fmulp
			frndint
			fistp	x1
			add		x1,		ecx	
//			xchg	eax,	x1
//			add		eax,	ecx

//		x2 = p1.x + f3d_fround(sy_py * k12);
			fld     k12
			fmulp
			frndint
			fistp	x2
			add		x2,		ecx
//			xchg	edx,	x2
//			add		edx,	ecx

//			f3d_HLine(x1, x2, sy, color);
//			push	color
//			push	sy
//			push	edx
//			push	eax
//			call	f3d_HLine
//			add		esp,	0x10
		}

		f3d_HLine(x1, x2, sy, color);

	}  // for

out_2_part:

// Верхняя часть

	if ((p3.y - p2.y) == 0) return F3D_NO_ERROR; // нет верхей половины
	k23 = (double)(signed(p3.x) - signed(p2.x)) / (double)(signed(p3.y) - signed(p2.y));

	sy_s = f3d_max(p2.y+1, 0);
	if (sy_s >= BuffHeight) return F3D_NO_ERROR; // верхняя половина не видна
	sy_f = f3d_min(p3.y, BuffHeight-1);

	sy_py = (sy_s - p1.y);
	sy_py2 = (sy_s - p2.y);

	for (sy = sy_s; sy <= sy_f; sy++, sy_py++, sy_py2++)
	{
		asm
		{
//		x1 = p1.x + f3d_fround(sy_py * k13);
			fld		sy_py
			fld		k13
			fmulp
			frndint
			fistp	x1
			mov		eax,	p1x
			add		x1,		eax

//		x2 = p2.x + f3d_fround(sy_py2 * k23);
			fld		sy_py2
			fld		k23
			fmulp
			frndint
			fistp	x2
			mov		eax,	p2x
			add		x2,		eax
		}

		f3d_HLine(x1, x2, sy, color);

	} // for

	return F3D_NO_ERROR;
}

//######################################################################################################################
TRIPLE Triple(double x, double y, double z)
{
    TRIPLE t;
    t.x = x;
    t.y = y;
    t.z = z;
    return t;
}

//######################################################################################################################
U32 f3d_ZbufInit()
{
	if (Zbuf)
	{
		return F3D_CRAZY_CALL;
	}

	Zbuf = (double*)LocalAlloc(LMEM_FIXED, ColorBuffSize*sizeof(double));
    if (!Zbuf) return false;

    f3d_ZbufClear();
    return F3D_NO_ERROR;
}

//######################################################################################################################
U32 f3d_ZbufDone()
{
	if (!Zbuf)
	{
		return F3D_CRAZY_CALL;
    }

    LocalFree(Zbuf);
    Zbuf = NULL;

    return F3D_NO_ERROR;
}

//######################################################################################################################
U32 f3d_ZbufClear()
{
	if (!Zbuf)
	{
		return F3D_CRAZY_CALL;
	}

	double *p=Zbuf;

	for (U32 i=0; i<ColorBuffSize; i++)
		*p++ = MAXDOUBLE;

	return F3D_NO_ERROR;
}

//######################################################################################################################
U32 f3d_SetPixel3D(S32 x, S32 y, double z, COLOR color)
{
	if ((!Zbuf) || (!lpColor))
	{
		return F3D_CRAZY_CALL;
    }

	if ((x<0) || (y<0) || (x>=BuffWidth) || (y>=BuffHeight)) return F3D_DATA_ERROR;

	U32 i = ((U32)y*(U32)BuffWidth+(U32)x);
	double *p = &Zbuf[i];

	if (!Z_BUF_SCALE_USED) z *= Z_BUF_SCALE; // увеличиваю точность

	if ((*p < z)) return F3D_NO_ERROR;
	*p = z;

	asm
	{
		mov		eax,	y
		mov		ecx,	x
		xor		edx,	edx
		mul		ColorBuffLine

		mov		edx,	ecx
		lea		edx,	[edx+2*edx]		// edx = x*3

		add		eax,	edx
		add		eax,	lpColor			// eax = lpColor + (ColorBuffLine*y) + (x*3)

		mov		ecx,	color

		mov		word ptr [eax],	cx
		inc		eax
		inc		eax
		shr		ecx,	16
		mov		byte ptr [eax], cl

		xor		eax,	eax
		inc		eax
		mov		i,		eax		// return eax = F3D_NO_ERROR

//		pop		ebp
//		ret
	}

	return i;
}

//######################################################################################################################
U32 f3d_HLine3D(S32 x1, S32 x2, double z1, double z2, S32 y, COLOR color)
{
	if ((!Zbuf) || (!lpColor))
	{
		return F3D_CRAZY_CALL;
	}

    // проверки на банальную видимость и порядок точек

	if ((y<0) || (y>=BuffHeight)) return F3D_DATA_ERROR;

	if (x1 > x2)
	{
		asm
		{
//		S32 x = x1; x1 = x2; x2 = x;
			mov		eax,	x1
			xchg	eax,	x2
			xchg	eax,	x1

//		double z = z1; z1 = z2; z2 = z;
			fld		qword ptr [z1]
			fld		qword ptr [z2]
			fstp	qword ptr [z1]
			fstp	qword ptr [z2]
		}
	}
#ifdef CALL_OPTIMIZE
	else
	if (x1 == x2)
	{
		return f3d_SetPixel3D(x1, y, z1, color);
	}
#endif // CALL_OPTIMIZE

	// есть ли смысл обрабатывать в оптимизации вариант с одним пикселем в нуле (x1<=0, x2==0)?

	if ((x1 >= BuffWidth) || (x2 < 0)) return F3D_DATA_ERROR;

#ifndef CALL_OPTIMIZE
	if (x1 == x2) return F3D_DATA_ERROR;
#endif // !CALL_OPTIMIZE

	if (!Z_BUF_SCALE_USED)
	{
		z1 *= Z_BUF_SCALE; // увеличиваю точность
		z2 *= Z_BUF_SCALE;
	}

	double dZ = (z2-z1)/(x2-x1);

	if (x1 < 0)
	{
		z1 -= x1*dZ;
		x1 = 0;
	}

	if (x2 >= BuffWidth)
	{
		z2 += (double)(BuffWidth-x2)*dZ; // "±1" ???
		x2 = BuffWidth-1;
	}

	// "точка входа" в массивы цвета и глубины
	U32 d = (U32)y*(U32)BuffWidth + (U32)x1;
	BYTE *pc = &lpColor[(d << 1)+d]; // d*3 {24bit color}
	double *pz = &Zbuf[d]; // d*1 (double Zbuffer)
	BYTE *c;

	// вывод линии
	while (x1<=x2)
	{
		if (*pz >= z1)
		{
			c = (BYTE*)&color;
			*(pc) = *c++;
			*(pc+1) = *c++;
			*(pc+2) = *c;

			*pz = z1;
		}

		pz++;
		pc += 3;

		z1 += dZ;
		x1++;
	}

	return F3D_NO_ERROR;
}

//######################################################################################################################
void swap(TRIPLE &v1, TRIPLE &v2)
{
	TRIPLE tmp = v1; v1 = v2; v2 = tmp;
}

//######################################################################################################################
U32 f3d_Triangle3D(TRIPLE &p1, TRIPLE &p2, TRIPLE &p3, COLOR color)
// ### модификация двумерного варианта (+ учитываю z-координату при выводе горизонтальных отрезков) ###
// НЕОБХОДИМО заменить тангенсы углов наклона на интерполяцию !!!!! 
{
	if ((!Zbuf) || (!lpColor))
	{
		return F3D_CRAZY_CALL;
	}

    // сортировка вершин
	if (p1.y > p2.y)
		// поменять местами вершины 1 и 2
		swap(p1, p2);
	if (p2.y > p3.y)
		// поменять местами вершины 2 и 3
		swap(p2, p3);
	if (p1.y > p2.y)
		// поменять местами вершины 1 и 2
		swap(p1, p2);

	S32 x1, x2, x3, sy, p1x, p2x, sy_s, sy_f, sy_py, sy_py2;

	double k12, k13, k23, tmp;
	double z12, z13, z23, dz12, dz13, dz23;

	if (!Z_BUF_SCALE_USED)
	{
		p1.z *= Z_BUF_SCALE;
		p2.z *= Z_BUF_SCALE;
		p3.z *= Z_BUF_SCALE;
	}

	tmp = p3.y - p1.y + 1;
	if (tmp == 0) return F3D_DATA_ERROR; // вырожденный треугольник
	dz13 = (p3.z - p1.z + 1) / tmp;
	k13 = (p3.x - p1.x + 1) / tmp;

    // указываю, что необходимо округлять числа в меньшую сторону
	short int mem16;
	asm
	{
		fstcw	mem16
		and		mem16,	1111001111111111b
		or		mem16,	1111011111111111b
		fldcw	mem16
	}

	p1x = f3d_fround(p1.x);
	p2x = f3d_fround(p2.x);

	z13 = p1.z;

//    goto out_2_part;

	tmp = p2.y - p1.y + 1;
	if (!tmp || (p2.y < 0)) goto out_2_part; // нет нижней половины, либо она не видна

// Нижняя часть

	k12 = (p2.x - p1.x+1) / tmp;
	dz12 = (p2.z - p1.z) / tmp;

	sy_py = 0;

//	if (p1.y > 0) sy_s = p1.y; else sy_s = 0;
	sy_s = /*f3d_max(*/p1.y/*, 0)*/;
// {v. 1.01>}
//	sy_f = f3d_min(p2.y, BuffHeight-1);
	sy_f = /*f3d_min(*/p2.y/*, BuffHeight-1)*/;
// {v. 1.01<}

	z12 = p1.z;

	Z_BUF_SCALE_USED = true;

	for (sy = sy_s; sy <= sy_f; sy++, sy_py++, z12 += dz12, z13 += dz13)
	{
		if (sy < 0) continue;
		if (sy >= BuffHeight)
		{
			Z_BUF_SCALE_USED = false;
			return F3D_NO_ERROR;
		}

		asm
		{
			mov		ecx,	p1x

//		x1 = p1.x + f3d_fround(sy_py * k13);
			fild	sy_py
			fld		st(0)
			fld		k13
			fmulp
			frndint
			fistp	x1
			add		x1,		ecx

//		x2 = p1.x + f3d_fround(sy_py * k12);
			fld		k12
			fmulp
			frndint
			fistp	x2
			add		x2,		ecx
		}

		f3d_HLine3D(x1, x2, z13, z12, sy, /*0xFFFFFF-*/color);

	}  // for

	Z_BUF_SCALE_USED = false;


out_2_part:

// Верхняя часть

	tmp = p3.y - p2.y + 1;
	if (!tmp) return F3D_NO_ERROR; // нет верхей половины
	k23 = (p3.x - p2.x + 1) / tmp;
	z23 = p2.z;
	dz23 = (p3.z - p2.z + 1) / tmp;

// {v. 1.01>}
//	sy_s = f3d_max(p2.y+1, 0);
//	if (p2.y > 0) sy_s = p2.y; else sy_s = 0;
	sy_s = /*f3d_max(*/p2.y+1/*, 0)*/;
// {v. 1.01<}
	if (sy_s >= BuffHeight) return F3D_NO_ERROR; // верхняя половина не видна
	sy_f = /*f3d_min(*/p3.y/*, BuffHeight-1)*/;

//	sy_py = (sy_s - f3d_fround(p1.y));
	sy_py2 = 0/*(sy_s - f3d_fround(p2.y))*/;

	Z_BUF_SCALE_USED = true;

	for (sy = sy_s; sy <= sy_f; sy++, sy_py++, sy_py2++, z13 += dz13, z23 += dz23)
	{
		if (sy < 0) continue;
		if (sy >= BuffHeight)
		{
			Z_BUF_SCALE_USED = false;
			return F3D_NO_ERROR;
		}
		
		asm
		{
//		x1 = p1.x + f3d_fround(sy_py * k13);
			fild	sy_py
			fld		k13
			fmulp
			frndint
			fistp	x1
			mov		eax,	p1x
			add		x1,		eax

//		x2 = p2.x + f3d_fround(sy_py2 * k23);
			fild	sy_py2
			fld		k23
			fmulp
			frndint
			fistp	x2
			mov		eax,	p2x
			add		x2,		eax
		}

		f3d_HLine3D(x1, x2, z13, z23, sy, color);

	} // for

	Z_BUF_SCALE_USED = false;

	return F3D_NO_ERROR;
}