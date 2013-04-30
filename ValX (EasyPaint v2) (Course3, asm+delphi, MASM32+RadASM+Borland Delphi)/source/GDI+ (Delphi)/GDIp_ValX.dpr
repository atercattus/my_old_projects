library GDIp_ValX;

uses Windows;

const WINGDIPDLL = 'gdiplus.dll';

type
    TStatus = (
        GDIp_Ok,
        GDIp_GenericError,
        GDIp_InvalidParameter,
        GDIp_OutOfMemory,
        GDIp_ObjectBusy,
        GDIp_InsufficientBuffer,
        GDIp_NotImplemented,
        GDIp_Win32Error,
        GDIp_WrongState,
        GDIp_Aborted,
        GDIp_FileNotFound,
        GDIp_ValueOverflow,
        GDIp_AccessDenied,
        GDIp_UnknownImageFormat,
        GDIp_FontFamilyNotFound,
        GDIp_FontStyleNotFound,
        GDIp_NotTrueTypeFont,
        GDIp_UnsupportedGdiplusVersion,
        GDIp_GdiplusNotInitialized,
        GDIp_PropertyNotFound,
        GDIp_PropertyNotSupported
    );

    GPIMAGE = pointer;
    GPGRAPHICS = pointer;

    GdiplusStartupInput = packed record
        GdiplusVersion          : Cardinal;       // Must be 1
        DebugEventCallback      : pointer; // Ignored on free builds
        SuppressBackgroundThread: BOOL;           // FALSE unless you're prepared to call
                                                  // the hook/unhook functions properly
        SuppressExternalCodecs  : BOOL;           // FALSE unless you want GDI+ only to use
    end;                                        // its internal image codecs.
    TGdiplusStartupInput = GdiplusStartupInput;
    PGdiplusStartupInput = ^TGdiplusStartupInput;

    TPIXELFORMAT = integer;

    TImageCodecInfo = packed record
        Clsid             : TGUID;
        FormatID          : TGUID;
        CodecName         : PWCHAR;
        DllName           : PWCHAR;
        FormatDescription : PWCHAR;
        FilenameExtension : PWCHAR;
        MimeType          : PWCHAR;
        Flags             : DWORD;
        Version           : DWORD;
        SigCount          : DWORD;
        SigSize           : DWORD;
        SigPattern        : PBYTE;
        SigMask           : PBYTE;
    end;
    PImageCodecInfo = ^TImageCodecInfo;

    TEncoderParameter = packed record
        Guid           : TGUID;   // GUID of the parameter
        NumberOfValues : ULONG;   // Number of the parameter values
        Type_          : ULONG;   // Value type, like ValueTypeLONG  etc.
        Value          : Pointer; // A pointer to the parameter values
    end;
    PEncoderParameter = ^TEncoderParameter;

    TEncoderParameters = packed record
        Count     : UINT;               // Number of parameters in this structure
        Parameter : array[0..0] of TEncoderParameter;  // Parameter values
    end;
    PEncoderParameters = ^TEncoderParameters;    

    TGdiplusStartup = function (out token: ULONG; input: PGdiplusStartupInput; output: pointer): TStatus; stdcall;
    TGdiplusShutdown = procedure (token: ULONG); stdcall;
    TGdipLoadImageFromFile = function (filename: PWCHAR; out image: GPIMAGE): TStatus; stdcall;
    TGdipDisposeImage = function (image: GPIMAGE): TStatus; stdcall;
	TGdipGetImageWidth = function (image: GPIMAGE; var width: UINT): TStatus; stdcall;
    TGdipGetImageHeight = function (image: GPIMAGE; var height: UINT): TStatus; stdcall;
    TGdipGetImagePixelFormat = function (image: GPIMAGE; out format: TPIXELFORMAT): TStatus; stdcall;
    TGdipCreateFromHDC = function (hdc: HDC; out graphics: GPGRAPHICS): TStatus; stdcall;
    TGdipDrawImageRectI = function (graphics: GPGRAPHICS; image: GPIMAGE; x: Integer; y: Integer; width: Integer; height: Integer): TStatus; stdcall;

    TGdipGetImageEncodersSize = function (out numEncoders: UINT; out size: UINT): TStatus; stdcall;
    TGdipGetImageEncoders = function (numEncoders: UINT; size: UINT; encoders: PIMAGECODECINFO): TStatus; stdcall;
    TGdipSaveImageToFile = function (image: GPIMAGE; filename: PWCHAR; clsidEncoder: PGUID; encoderParams: PEncoderParameters): TStatus; stdcall;
    TGdipCreateBitmapFromHBITMAP = function (hbm: HBITMAP; hpal: HPALETTE; out bitmap: pointer): TStatus; stdcall;

var
	GdiplusStartup : TGdiplusStartup;
    GdiplusShutdown : TGdiplusShutdown;
    GdipLoadImageFromFile : TGdipLoadImageFromFile;
    GdipDisposeImage : TGdipDisposeImage;
    GdipGetImageWidth : TGdipGetImageWidth;
    GdipGetImageHeight : TGdipGetImageHeight;
    GdipGetImagePixelFormat : TGdipGetImagePixelFormat;
    GdipCreateFromHDC : TGdipCreateFromHDC;
    GdipDrawImageRectI : TGdipDrawImageRectI;

    GdipGetImageEncodersSize : TGdipGetImageEncodersSize;
    GdipGetImageEncoders : TGdipGetImageEncoders;
    GdipSaveImageToFile : TGdipSaveImageToFile;
    GdipCreateBitmapFromHBITMAP : TGdipCreateBitmapFromHBITMAP;

    //##################################################################################################################
	function GDIp_LoadImage(path : PChar) : HBITMAP; stdcall; export;
    var
    	hDLL			: DWORD;
        res				: TStatus;
        StartupInput	: TGDIPlusStartupInput;
        gdiplusToken	: ULONG;
        image			: pointer;
        W, H			: DWORD;
        PixelFormat		: integer;
        graphics		: pointer;
        BPP				: DWORD;
        OldBMP			: HBITMAP;

		// собственный контекст
		DC				: hDC;
    	// собственный bitmap (это и есть результат)
		Bitmap			: HBITMAP;		
		// параметры bitmap'а
		bitmap_info		: BITMAPINFO;
		// собственный буфер цвета
		lpColor			: pointer;
	begin
    	hDLL := LoadLibrary(WINGDIPDLL);
        if (hDLL=0) then
        begin
			Result := 0;
            exit;
        end;

        GdiplusStartup := GetProcAddress(hDLL, 'GdiplusStartup');
        GdiplusShutdown := GetProcAddress(hDLL, 'GdiplusShutdown');
        GdipLoadImageFromFile := GetProcAddress(hDLL, 'GdipLoadImageFromFile');
        GdipDisposeImage := GetProcAddress(hDll, 'GdipDisposeImage');
        GdipGetImageWidth := GetProcAddress(hDLL, 'GdipGetImageWidth');
        GdipGetImageHeight := GetProcAddress(hDLL, 'GdipGetImageHeight');
        GdipGetImagePixelFormat := GetProcAddress(hDLL, 'GdipGetImagePixelFormat');
        GdipCreateFromHDC := GetProcAddress(hDLL, 'GdipCreateFromHDC');
        GdipDrawImageRectI := GetProcAddress(hDLL, 'GdipDrawImageRectI');

        if	(@GdiplusStartup=nil) or
        	(@GdiplusShutdown=nil) or
            (@GdipLoadImageFromFile=nil) or
            (@GdipDisposeImage=nil) or
            (@GdipGetImageWidth=nil) or
            (@GdipGetImageHeight=nil) or
            (@GdipGetImagePixelFormat=nil) or
            (@GdipCreateFromHDC=nil) or
            (@GdipDrawImageRectI=nil) then
        begin
            FreeLibrary(hDLL);
        	Result := DWORD(-1);
            exit;
        end;

        StartupInput.DebugEventCallback := nil;
        StartupInput.SuppressBackgroundThread := false;
        StartupInput.SuppressExternalCodecs   := false;
        StartupInput.GdiplusVersion := 1;
        // Initialize GDI+
        res := GdiplusStartup(gdiplusToken, @StartupInput, nil);
        if (res <> GDIp_Ok) then
        begin
            FreeLibrary(hDLL);
            Result := DWORD(-2);
            exit; 
        end;

        // Загрузка изображения
        image := nil;
		res := GdipLoadImageFromFile(PWideChar(WideString(path)), image);
        if (res <> GDIp_Ok) then
        begin
			GdiplusShutdown(gdiplusToken);
            FreeLibrary(hDLL);
            Result := DWORD(-3);
            exit;
        end;

		if (	(GdipGetImageWidth(image, W) <> GDIp_Ok) or
        		(GdipGetImageHeight(image, H) <> GDIp_Ok) or
                (GdipGetImagePixelFormat(image, PixelFormat) <> GDIp_Ok) or
                (W=0) or (H=0)) then
        begin
        	GdipDisposeImage(image);
			GdiplusShutdown(gdiplusToken);
            FreeLibrary(hDLL);
            Result := DWORD(-4);
            exit;
        end;

        case (PixelFormat and $0F) of
			1		: BPP := 1;
            2		: BPP := 4;
            3		: BPP := 8;
            4..7	: BPP := 16;
            8		: BPP := 24;
            9..11	: BPP := 32;
            else
            	begin
                	GdipDisposeImage(image);
                    GdiplusShutdown(gdiplusToken);
                    FreeLibrary(hDLL);
                    Result := DWORD(-5);
                    exit;
                end;
        end;

        // Создание битмапа и контекста
        DC := CreateCompatibleDC(GetDC(0));
        if (DC=0) then
        begin
        	GdipDisposeImage(image);
			GdiplusShutdown(gdiplusToken);
            FreeLibrary(hDLL);
            Result := DWORD(-6);
            exit;
        end;

        // формирую BITMAPINFO
        ZeroMemory(@bitmap_info, sizeof(BITMAPINFO));

        bitmap_info.bmiHeader.biSize := sizeof(BITMAPINFOHEADER);
        bitmap_info.bmiHeader.biWidth := W;
        bitmap_info.bmiHeader.biHeight := H;
        bitmap_info.bmiHeader.biPlanes := 1;
        bitmap_info.bmiHeader.biBitCount := BPP;
        bitmap_info.bmiHeader.biCompression := BI_RGB;

        // создаю буфер цвета
        Bitmap := CreateDIBSection(DC, bitmap_info, DIB_RGB_COLORS, lpColor, 0, 0);
        if (lpColor=nil) or (Bitmap=0) then
		begin
            DeleteDC(DC);
            GdipDisposeImage(image);
			GdiplusShutdown(gdiplusToken);
            FreeLibrary(hDLL);
            Result := DWORD(-7);
            exit;
        end;

        OldBMP := SelectObject(DC, Bitmap);

        res := GdipCreateFromHDC(DC, graphics);
        if (res <> GDIp_Ok) then
        begin
        	DeleteObject(Bitmap);
        	DeleteDC(DC);
            GdipDisposeImage(image);
			GdiplusShutdown(gdiplusToken);
            FreeLibrary(hDLL);
            Result := DWORD(-8);
            exit;
        end;        

        res := GdipDrawImageRectI(graphics, image, 0, 0, W, H);
        if (res <> GDIp_Ok) then
        begin
        	DeleteObject(Bitmap);
            GdipDisposeImage(image);
			GdiplusShutdown(gdiplusToken);
            FreeLibrary(hDLL);
            Result := DWORD(-9);
            exit;
        end;

		SelectObject(DC, OldBMP);

		DeleteDC(DC);

		// Close GDI+
        GdipDisposeImage(image);
        GdiplusShutdown(gdiplusToken);

        FreeLibrary(hDLL);

        Result := Bitmap;
    end;

	//##################################################################################################################
    function GetEncoderClsid(format : string; out pClsid : TGUID): integer; stdcall;
    type
   		ArrImgInf = array of TImageCodecInfo;

    var
        num, size: UINT;
        j : longint;
        ImageCodecInfo : PImageCodecInfo;
        s : string;
    begin
        num  := 0; // number of image encoders
        size := 0; // size of the image encoder array in bytes
        Result := -1;

        GdipGetImageEncodersSize(num, size);
        if (size = 0) then exit;

        GetMem(ImageCodecInfo, size);
        if(ImageCodecInfo = nil) then exit;

        GdipGetImageEncoders(num, size, ImageCodecInfo);

        for j := 0 to num - 1 do
        begin
        	s := WideString(ArrImgInf(ImageCodecInfo)[j].MimeType);
            if(s = format) then
            begin
                pClsid := ArrImgInf(ImageCodecInfo)[j].Clsid;
                Result := j;  // Success
            end;
        end;
        FreeMem(ImageCodecInfo, size);
    end;

	//##################################################################################################################
    function GDIp_SaveImage(B : HBITMAP; path : PChar; encoder : PChar) : DWORD; stdcall; export;
    var
    	hDLL			: DWORD;
        res				: TStatus;
        StartupInput	: TGDIPlusStartupInput;
        gdiplusToken	: ULONG;
        image			: pointer;
        encoderClsid	: TGUID;

	begin
    	hDLL := LoadLibrary(WINGDIPDLL);
        if (hDLL=0) then
        begin
			Result := 1;
            exit;
        end;

        GdiplusStartup := GetProcAddress(hDLL, 'GdiplusStartup');
        GdiplusShutdown := GetProcAddress(hDLL, 'GdiplusShutdown');
        GdipGetImageEncodersSize := GetProcAddress(hDLL, 'GdipGetImageEncodersSize');
        GdipGetImageEncoders := GetProcAddress(hDLL, 'GdipGetImageEncoders');
        GdipSaveImageToFile := GetProcAddress(hDLL, 'GdipSaveImageToFile');
        GdipCreateBitmapFromHBITMAP := GetProcAddress(hDLL, 'GdipCreateBitmapFromHBITMAP');
        GdipDisposeImage := GetProcAddress(hDLL, 'GdipDisposeImage');

        if	(@GdiplusStartup=nil) or
        	(@GdiplusShutdown=nil) or
            (@GdipGetImageEncodersSize=nil) or
            (@GdipGetImageEncoders=nil) or
            (@GdipSaveImageToFile=nil) or
            (@GdipCreateBitmapFromHBITMAP=nil) or
            (@GdipDisposeImage=nil) then
        begin
            FreeLibrary(hDLL);
        	Result := 2;
            exit;
        end;

        StartupInput.DebugEventCallback := nil;
        StartupInput.SuppressBackgroundThread := false;
        StartupInput.SuppressExternalCodecs   := false;
        StartupInput.GdiplusVersion := 1;
        // Initialize GDI+
        res := GdiplusStartup(gdiplusToken, @StartupInput, nil);
        if (res <> GDIp_Ok) then
        begin
            FreeLibrary(hDLL);
            Result := 3;
            exit;
        end;

        // Создаю GDI+ bitmap из HBITMAP 
        res := GdipCreateBitmapFromHBITMAP(B, 0, image); // палитру можно через GetObject HPALETTE
        if (res <> GDIp_Ok) then
        begin
        	GdiplusShutdown(gdiplusToken);
            FreeLibrary(hDLL);
            Result := 4;
            exit;
        end;

        if (GetEncoderClsid(encoder, encoderClsid) = -1) then
        begin
        	GdipDisposeImage(image);
        	GdiplusShutdown(gdiplusToken);
            FreeLibrary(hDLL);
            Result := 5;
            exit;
        end;

        res := GdipSaveImageToFile(image, PWideChar(WideString(path)), @encoderClsid, nil);
        if (res <> GDIp_Ok) then
        begin
        	GdipDisposeImage(image);
        	GdiplusShutdown(gdiplusToken);
            FreeLibrary(hDLL);
            Result := 6;
            exit;
        end;

		// Close GDI+
        GdipDisposeImage(image);
        GdiplusShutdown(gdiplusToken);

        FreeLibrary(hDLL);

        Result := 0;
    end;

exports
	GDIp_LoadImage,
    GDIp_SaveImage;

end.
