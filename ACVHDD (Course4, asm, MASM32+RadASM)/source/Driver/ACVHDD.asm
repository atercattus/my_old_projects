include consts.inc
include AES\AESCrypt.inc
include dispatch.asm

.code

AC_OpenFileDrive proc uses esi edi \
	pDeviceObject	: PDEVICE_OBJECT, \
	pIrp			: PIRP
	
LOCAL pusFileOpenName	: UNICODE_STRING
LOCAL status			: NTSTATUS
LOCAL iosb				: IO_STATUS_BLOCK
LOCAL oa				: OBJECT_ATTRIBUTES
LOCAL file_standart		: FILE_STANDARD_INFORMATION
LOCAL file_alignment	: FILE_ALIGNMENT_INFORMATION

	mov		status, STATUS_UNSUCCESSFUL
	assume esi : ptr _IRP
	mov		esi, pIrp
	
	; Проверим валидность буфера
	mov		eax, [esi].AssociatedIrp.SystemBuffer
	_try
		mov al, byte ptr [eax]
	_finally
	.if !Exception
		mov		status, STATUS_INVALID_PARAMETER
		jmp		ACOFD_exit
	.endif

	invoke	RtlInitUnicodeString, addr pusFileOpenName, [esi].AssociatedIrp.SystemBuffer
	lea		ecx, oa
	InitializeObjectAttributes	ecx, addr pusFileOpenName, OBJ_CASE_INSENSITIVE + OBJ_KERNEL_HANDLE, NULL, NULL
	invoke	ZwCreateFile,	addr hDiskFile, FILE_READ_DATA + FILE_WRITE_DATA, addr oa,\
							addr iosb, 0, FILE_ATTRIBUTE_NORMAL, 0, FILE_OPEN,\
							FILE_SYNCHRONOUS_IO_NONALERT or FILE_NON_DIRECTORY_FILE or FILE_RANDOM_ACCESS or\
							FILE_NO_INTERMEDIATE_BUFFERING, NULL, 0
	mov		status, eax

	.if	(eax != STATUS_SUCCESS)
		invoke	ZwCreateFile,	addr hDiskFile, FILE_READ_DATA, addr oa,\
		  						addr iosb, 0, FILE_ATTRIBUTE_NORMAL, 0, FILE_OPEN,\
		  						FILE_SYNCHRONOUS_IO_NONALERT or FILE_NON_DIRECTORY_FILE or FILE_RANDOM_ACCESS or\
		  						FILE_NO_INTERMEDIATE_BUFFERING, NULL, 0
		mov		status, eax
		.if		(eax != STATUS_SUCCESS)
			jmp ACOFD_exit
		.endif
	.endif

	invoke	ZwQueryInformationFile,	hDiskFile, addr status, addr file_alignment,\
									sizeof(FILE_ALIGNMENT_INFORMATION), FileAlignmentInformation
	lea		edi, file_alignment
	assume edi : ptr FILE_ALIGNMENT_INFORMATION
	mov		eax, [edi].AlignmentRequirement
	mov		edi, pDeviceObject
	assume edi : ptr DEVICE_OBJECT
	mov		[edi].AlignmentRequirement, eax
	assume edi : ptr nothing

	invoke	ZwQueryInformationFile,	hDiskFile, addr status, addr file_standart,\
									sizeof(FILE_STANDARD_INFORMATION), FileStandardInformation
	lea		edi, file_standart
	assume edi:ptr FILE_STANDARD_INFORMATION
	mov		eax, [edi].EndOfFile.LowPart
	and		eax, 0fff00000h; Округляем до 1 мб
	.if	(eax==0)
		mov	status, STATUS_DRIVER_INTERNAL_ERROR
		jmp	ACOFD_exit
	.endif
	mov		CreateParams.nDiskSize.LowPart,		eax
	mov		CreateParams.nDiskSize.HighPart,	0

	mov		ecx, BYTES_PER_SECTOR
	xor		edx, edx
	div		ecx
	xor		edx, edx
	mov		ecx, SECTORS_PER_TRACK
	div		ecx
	xor		edx, edx
	mov		ecx, TRACKS_PER_CYLINDER
	div		ecx
;Cylinders = nDiskSize/512/32/2
	mov		CreateParams.Cylinders.LowPart, eax
	mov		CreateParams.Cylinders.HighPart, 0
	mov		CreateParams.MediaType, FixedMedia
	mov		CreateParams.TracksPerCylinder, TRACKS_PER_CYLINDER
	mov		CreateParams.SectorsPerTrack, SECTORS_PER_TRACK
	mov		CreateParams.BytesPerSector, BYTES_PER_SECTOR
	mov		CreateParams.HiddenSectors, HIDDEN_SECTORS

ACOFD_exit:
	mov eax, status
	
	ret
AC_OpenFileDrive endp

AC_CloseFileDrive proc uses edi \
	pDeviceObject	: PDEVICE_OBJECT,\
	pIrp			: PIRP

;LOCAL	iosb			: IO_STATUS_BLOCK
;LOCAL	FileBasicInfo	: FILE_BASIC_INFORMATION

	.if	(pMemory != NULL)
		invoke	ExFreePool, pMemory
		and		pMemory, NULL
	.endif
	
	; Завершение работы системы шифрования
	invoke	DoneAES

	; Закрываем файл
	invoke	ZwClose, hDiskFile
	mov		hDiskFile, FALSE

	ret
AC_CloseFileDrive endp

AC_DriverUnload proc \
	pDriverObject : PDRIVER_OBJECT

	.if	(hDiskFile != NULL)
		invoke	ZwClose, hDiskFile
	.endif
	
	mov		eax, pDriverObject
	assume eax:ptr DRIVER_OBJECT
	invoke	IoDeleteDevice, [eax].DeviceObject
	assume eax:ptr nothing
	
	.if	(pMemory != NULL)
		invoke	RtlZeroMemory, pMemory, BufferLength
		invoke	ExFreePool, pMemory
	.endif
	
	; Завершение работы системы шифрования
	invoke	DoneAES
	
	mov eax, STATUS_SUCCESS

	ret
AC_DriverUnload endp

.code INIT

AC_DriverEntry proc uses esi esi \
	pDriverObject	: PDRIVER_OBJECT, \
	pusRegistryPath	: PUNICODE_STRING

LOCAL	status			: NTSTATUS
LOCAL	pDeviceObject	: PDEVICE_OBJECT

	mov		status,			STATUS_DEVICE_CONFIGURATION_ERROR
	and		hDiskFile,		NULL
	and		nOpenCount,		0
	and		BufferLength,	0
	and		pMemory,		0
	
	invoke	IoCreateDevice,	pDriverObject, 0, addr usDeviceName,\
							FILE_DEVICE_DISK, 0, FALSE, addr pDeviceObject
	.if (eax == STATUS_SUCCESS)
		mov		status,	eax
		mov		eax,	pDeviceObject
		or		(DEVICE_OBJECT PTR [eax]).Flags, DO_DIRECT_IO
		mov		eax,	pDriverObject
		assume eax:ptr DRIVER_OBJECT
		mov		[eax].MajorFunction[IRP_MJ_CREATE*(sizeof PVOID)],			offset AC_DispatchCreate
		mov		[eax].MajorFunction[IRP_MJ_CLOSE*(sizeof PVOID)],			offset AC_DispatchClose
		mov		[eax].MajorFunction[IRP_MJ_READ*(sizeof PVOID)],			offset AC_DispatchReadWrite
		mov		[eax].MajorFunction[IRP_MJ_WRITE*(sizeof PVOID)],			offset AC_DispatchReadWrite
		mov		[eax].MajorFunction[IRP_MJ_DEVICE_CONTROL*(sizeof PVOID)],	offset AC_DispatchControl
		mov		[eax].DriverUnload,											offset AC_DriverUnload
		assume eax:nothing
		invoke	RtlZeroMemory, addr CreateParams, sizeof(CREATE_PARAMS)

		MUTEX_INIT kDiskMutex
	.endif
	
	invoke RtlZeroMemory, addr CreateParams, sizeof CreateParams
	.if	(status != STATUS_SUCCESS)
		mov		eax,			pDriverObject
		assume eax:ptr DRIVER_OBJECT
		invoke	IoDeleteDevice,	[eax].DeviceObject
		assume eax:ptr nothing
		mov		pMemory,		0
	.endif
DE_exit:
	mov	eax, status

	ret
AC_DriverEntry endp

end AC_DriverEntry