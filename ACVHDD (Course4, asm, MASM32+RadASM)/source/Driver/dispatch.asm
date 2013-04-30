.code

MmGetSystemAddressForMdlSafe proc pMdl:PMDL, Priority:DWORD

	mov		eax, pMdl
	assume eax:ptr MDL
	
	.if ([eax].MdlFlags & (MDL_MAPPED_TO_SYSTEM_VA + MDL_SOURCE_IS_NONPAGED_POOL))
		mov	eax, [eax].MappedSystemVa
	.else
		invoke	MmMapLockedPagesSpecifyCache, pMdl, KernelMode, MmCached, NULL, FALSE, Priority
	.endif
	
	assume eax:nothing
	
	ret
MmGetSystemAddressForMdlSafe endp

AC_DispatchReadWrite proc uses esi edi ebx edx ecx\
	pDeviceObject	: PDEVICE_OBJECT,\
	pIrp			: PIRP

LOCAL	p			: PVOID
LOCAL	ByteOffset	: LARGE_INTEGER
LOCAL	iosb		: IO_STATUS_BLOCK
LOCAL	status		: NTSTATUS
LOCAL	modmod : DWORD

	mov	esi, pIrp
	assume esi:ptr _IRP

	mov		status, STATUS_NO_MEDIA_IN_DEVICE
	mov		[esi].IoStatus.Information, 0
	IoGetCurrentIrpStackLocation esi
	mov		edi, eax
	cmp		hDiskFile, NULL
	jz		DRWErrorExit
	assume edi:ptr IO_STACK_LOCATION
	.if	([edi].Parameters.Read._Length == 0)
		mov		[esi].IoStatus.Information, 0
		mov		status, STATUS_SUCCESS
		jmp		DRWErrorExit
	.endif

	.if (pMemory == NULL)
		invoke	ExAllocatePoolWithTag, NonPagedPool, [edi].Parameters.Read._Length, 'DprC'
		.if (eax != 0)
			mov		pMemory, eax
			push	[edi].Parameters.Read._Length
			pop		BufferLength
		.else
			mov		status, STATUS_INSUFFICIENT_RESOURCES
			jmp		DRWErrorExit
		.endif
	.endif
	
	mov		eax, BufferLength
	.if (eax < [edi].Parameters.Read._Length)
		MUTEX_WAIT kDiskMutex
		invoke	ExFreePool, pMemory
		mov		pMemory, NULL
		invoke	ExAllocatePoolWithTag, NonPagedPool, [edi].Parameters.Read._Length, 'DprC'
		.if	(eax != 0)
			mov		pMemory, eax
			push	[edi].Parameters.Read._Length
			pop		BufferLength
			MUTEX_RELEASE kDiskMutex
		.else
			mov		status, STATUS_INSUFFICIENT_RESOURCES
			MUTEX_RELEASE kDiskMutex
			jmp		DRWErrorExit
		.endif
	.endif

; Позиционируем положение указателя файла
	mov		eax, [edi].Parameters.Read.ByteOffset.HighPart
	mov		ByteOffset.HighPart, eax
	mov		eax, [edi].Parameters.Read.ByteOffset.LowPart
	mov		ByteOffset.LowPart, eax

	add		eax, [edi].Parameters.Read._Length

	.if	(eax <= CreateParams.nDiskSize.LowPart)
		; Get a system-space pointer to the user's buffer
		.if	([esi].MdlAddress != NULL)
			invoke	MmGetSystemAddressForMdlSafe, [esi].MdlAddress, NormalPagePriority
			.if	(eax != NULL)
				mov		p, eax
				; p содержит адрес буфера
				mov		eax, [edi].Parameters.Read._Length
				mov		[esi].IoStatus.Information, eax
				
				MUTEX_WAIT kDiskMutex
				
				.if	([edi].MajorFunction == IRP_MJ_READ)
					invoke	ZwReadFile,	hDiskFile, NULL, NULL, NULL, addr iosb, pMemory,\
										[edi].Parameters.Read._Length, addr ByteOffset, NULL
					mov status, eax
				; Расшифровка данных
;					mov		eax,	[edi].Parameters.Read._Length
;					and		eax,	3FFh
;					.IF (eax != 0)
;						mov		modmod,	eax
;						invoke	DbgPrint, $CTA0("\nEB: %8d\n"), modmod
;					.ENDIF
									
					mov		eax,	[edi].Parameters.Read._Length
					shr		eax,	4
					invoke	DecodeBuffer, pMemory, pMemory, eax
					invoke	memcpy, p, pMemory, [edi].Parameters.Read._Length
					
				.elseif ([edi].MajorFunction == IRP_MJ_WRITE)
					invoke	memcpy, pMemory, p, [edi].Parameters.Read._Length
				; Шифровка данных
;					mov		eax,	[edi].Parameters.Read._Length
;					and		eax,	3FFh
;					.IF (eax != 0)
;						mov		modmod,	eax
;						invoke	DbgPrint, $CTA0("\nDB: %8d\n"), modmod
;					.ENDIF
;					invoke	DbgPrint, $CTA0("\nDecodeBuffer: offs:%8X len:%8d\n"),
;									[edi].Parameters.Read.ByteOffset.LowPart,
;									[edi].Parameters.Read._Length
									
					mov		eax,	[edi].Parameters.Read._Length
					shr		eax,	4
					invoke	EncodeBuffer, pMemory, pMemory, eax
					invoke	ZwWriteFile,	hDiskFile, NULL, NULL, NULL, addr iosb, pMemory,\
											[edi].Parameters.Read._Length, addr ByteOffset, NULL
					mov		status, eax
				.endif
				
				MUTEX_RELEASE kDiskMutex
				
			.else
				mov	status, STATUS_INSUFFICIENT_RESOURCES
			.endif
		.endif
	.else
		mov status, STATUS_INVALID_PARAMETER
	.endif

	assume edi:nothing
DRWErrorExit:
	mov		eax, status
	mov		[esi].IoStatus.Status, eax
	fastcall IofCompleteRequest, esi, IO_NO_INCREMENT

	mov		eax, status

	ret
AC_DispatchReadWrite endp

AC_DispatchCreate proc \
	pDeviceObject	: PDEVICE_OBJECT,
	pIrp			: PIRP

	inc		nOpenCount

	mov		eax, pIrp
	mov		(_IRP PTR [eax]).IoStatus.Status, STATUS_SUCCESS
	and		(_IRP PTR [eax]).IoStatus.Information, 0

	fastcall IofCompleteRequest, pIrp, IO_NO_INCREMENT
	
	mov		eax, STATUS_SUCCESS

	ret
AC_DispatchCreate endp

AC_DispatchClose proc \
	pDeviceObject	: PDEVICE_OBJECT,
	pIrp			: PIRP
	
	.if	(nOpenCount != 0)
		dec	nOpenCount
	.endif

	mov		eax, pIrp
	mov		(_IRP PTR [eax]).IoStatus.Status, STATUS_SUCCESS
	and		(_IRP PTR [eax]).IoStatus.Information, 0

	fastcall IofCompleteRequest, pIrp, IO_NO_INCREMENT

	mov		eax, STATUS_SUCCESS
	
	ret
AC_DispatchClose endp

AC_LoadKeyAndInitAES proc \
	pDeviceObject	: PDEVICE_OBJECT,
	pIrp			: PIRP

LOCAL status : DWORD

	mov		status, STATUS_UNSUCCESSFUL
	mov		esi, pIrp
	assume esi : ptr _IRP
	
	; Проверю валидность буфера
	mov		Exception,	TRUE
	mov		eax, [esi].AssociatedIrp.SystemBuffer
	_try
		mov		al, byte ptr [eax]
	_finally
	.if (!Exception)
		mov		status, STATUS_INVALID_PARAMETER
		jmp		@@AC_LKI_exit
	.endif
	
	invoke	InitAES, [esi].AssociatedIrp.SystemBuffer
	; надо бы зачистить SystemBuffer для надежности...
	
	mov		status, STATUS_SUCCESS
	
@@AC_LKI_exit:
	mov		eax, status
	ret
AC_LoadKeyAndInitAES endp

AC_DispatchControl proc uses esi edi ebx edx ecx \
	pDeviceObject	: PDEVICE_OBJECT,
	pIrp			: PIRP

LOCAL	status : DWORD

	mov		esi, pIrp
	assume esi:ptr _IRP

	; Assume unsuccess
	mov		status, STATUS_UNSUCCESSFUL
	; We copy nothing
	mov		[esi].IoStatus.Information, 0

	IoGetCurrentIrpStackLocation esi
	mov		edi, eax
	assume edi:ptr IO_STACK_LOCATION

	mov		Exception, TRUE
	
	.if ([edi].Parameters.DeviceIoControl.IoControlCode == IOCTL_DISK_GET_DRIVE_GEOMETRY) || ([edi].Parameters.DeviceIoControl.IoControlCode == IOCTL_DISK_GET_MEDIA_TYPES)
		.if ([edi].Parameters.DeviceIoControl.OutputBufferLength >= sizeof DISK_GEOMETRY)
		; Запрос информации по дисковой геометрии
			mov		edi, [esi].AssociatedIrp.SystemBuffer
			; Проверим валидность буфера
			_try
				mov		al, byte ptr [edi]
			_finally
			.if (!Exception)
				mov		status, STATUS_INVALID_PARAMETER
				jmp		ACDC_exit
			.endif
			assume edi:ptr DISK_GEOMETRY
			mov		eax,						CreateParams.SectorsPerTrack
			;SectorsPerTrack
			mov		[edi].SectorsPerTrack,		eax
			;TracksPerCylinder
			mov		eax,						CreateParams.TracksPerCylinder
			mov		[edi].TracksPerCylinder,	eax
			mov		eax,						CreateParams.Cylinders.LowPart
			mov		ecx,						CreateParams.Cylinders.HighPart
			;Cylinders
			mov		[edi].Cylinders.LowPart,	eax
			mov		[edi].Cylinders.HighPart,	ecx
			mov		eax,						CreateParams.MediaType
			;MediaType
			mov		[edi].MediaType,			eax
			mov		eax,						CreateParams.BytesPerSector
			;BytesPerSector
			mov		[edi].BytesPerSector,		eax
			mov		status,	STATUS_SUCCESS
			mov		[esi].IoStatus.Information, sizeof(DISK_GEOMETRY)
		.else
			mov		status,	STATUS_BUFFER_TOO_SMALL
			mov		[esi].IoStatus.Information,	0
		.endif
		assume edi:ptr IO_STACK_LOCATION

	.elseif ([edi].Parameters.DeviceIoControl.IoControlCode == IOCTL_DISK_GET_PARTITION_INFO)
		.if ([edi].Parameters.DeviceIoControl.OutputBufferLength >= sizeof PARTITION_INFORMATION)
		; Запрос информации о PARTITION INFORMATION
			mov		edi, [esi].AssociatedIrp.SystemBuffer
			; Проверим валидность буфера
			_try
				mov		al, byte ptr [edi]
			_finally
			.if (!Exception)
				mov		status, STATUS_INVALID_PARAMETER
				jmp		ACDC_exit
			.endif
			
			assume edi:ptr PARTITION_INFORMATION
			; The offset in bytes on drive where the partition begins
			; StartingOffset
			mov		eax, CreateParams.BytesPerSector
			mul		CreateParams.HiddenSectors
			mov		[edi].StartingOffset.LowPart, eax
			mov		[edi].StartingOffset.HighPart, 0
			; The length in bytes of the partition
			mov		ecx, CreateParams.nDiskSize.LowPart
			; PartitionLength
			sub		ecx, eax
			mov		[edi].PartitionLength.LowPart, ecx
			mov		[edi].PartitionLength.HighPart, 0
			; The number of hidden sectors
			; HiddenSectors
			mov		eax, CreateParams.HiddenSectors
			mov		[edi].HiddenSectors, eax
			; Specifies the number of the partition
			; PartitionNumber
			mov		[edi].PartitionNumber, 0
			; Indicates the system-defined MBR partition type
			; PartitionType
			mov		byte ptr [edi].PartitionType, PARTITION_ENTRY_UNUSED
			; FALSE indicates that this partition is not bootable
			; BootIndicator
			mov		byte ptr [edi].BootIndicator, FALSE
			; RecognizedPartition
			; TRUE indicates that the system recognized the type of the partition
			mov		byte ptr [edi].RecognizedPartition, FALSE
			; RewritePartition
			; FALSE indicates that the partition information has not changed
			mov		byte ptr [edi].RewritePartition, FALSE
			;
			mov		status, STATUS_SUCCESS
			mov		[esi].IoStatus.Information, sizeof(PARTITION_INFORMATION)
			assume edi:ptr nothing
		.else
			mov		status, STATUS_INVALID_PARAMETER
			mov		[esi].IoStatus.Information, 0
		.endif
		assume edi:ptr IO_STACK_LOCATION

	.elseif ([edi].Parameters.DeviceIoControl.IoControlCode == IOCTL_DISK_GET_LENGTH_INFO)
		mov		edi, [esi].AssociatedIrp.SystemBuffer
		; Проверим валидность буфера
		_try
			mov		al, byte ptr [edi]
		_finally
		.if (!Exception)
			mov		status, STATUS_INVALID_PARAMETER
			jmp		ACDC_exit
		.endif
		;
		assume edi:ptr GET_LENGTH_INFORMATION
		mov		[edi]._Length.HighPart, 0
		mov		eax, CreateParams.nDiskSize.LowPart
		mov		[edi]._Length.LowPart, eax
		mov		status, STATUS_SUCCESS
		mov		[esi].IoStatus.Information, sizeof(GET_LENGTH_INFORMATION)
		assume edi:ptr nothing
		assume edi:ptr IO_STACK_LOCATION

	.elseif ([edi].Parameters.DeviceIoControl.IoControlCode == IOCTL_MOUNTDEV_QUERY_DEVICE_NAME)
		; Проверим валидность буфера
		mov		eax, [esi].AssociatedIrp.SystemBuffer
		_try
			mov		al, byte ptr [eax]
		_finally
		.if !Exception
			mov		status, STATUS_INVALID_PARAMETER
			jmp		ACDC_exit
		.endif
		
		lea		eax, usDeviceName
		assume eax:ptr UNICODE_STRING
		movzx	ecx, [eax].MaximumLength
		.if (ecx < [edi].Parameters.DeviceIoControl.OutputBufferLength)
			push	esi
			push	edi
			mov		[esi].IoStatus.Information, ecx
			add		[esi].IoStatus.Information, 2
			mov		edi, [esi].AssociatedIrp.SystemBuffer
			mov		esi, [eax].Buffer
			mov		word ptr [edi], cx
			add		edi, 2
			cld
			rep movsb
			assume eax:nothing
			pop		edi
			pop		esi
			mov		status, STATUS_SUCCESS
		.else
			mov		[esi].IoStatus.Information, 0
			mov		status, STATUS_BUFFER_TOO_SMALL
		.endif

	.elseif [edi].Parameters.DeviceIoControl.IoControlCode == IOCTL_UNMOUNT_DISK
		.if (nOpenCount == 0)
			invoke	AC_CloseFileDrive, pDeviceObject, pIrp
			mov	status, eax
		.else
			mov	status, STATUS_DEVICE_BUSY
		.endif
		mov		[esi].IoStatus.Information, 0

	.elseif [edi].Parameters.DeviceIoControl.IoControlCode == IOCTL_MOUNT_DISK
		invoke	AC_OpenFileDrive, pDeviceObject, pIrp
		mov		status, eax
		mov		[esi].IoStatus.Information, 0

	.elseif [edi].Parameters.DeviceIoControl.IoControlCode == IOCTL_SET_KEY_DISK
		; загрузка полученного ключа
		invoke	AC_LoadKeyAndInitAES, pDeviceObject, pIrp
		mov		status, STATUS_SUCCESS
		mov		[esi].IoStatus.Information, 0
		
	.elseif [edi].Parameters.DeviceIoControl.IoControlCode == IOCTL_DISK_SET_PARTITION_INFO
		mov		status, STATUS_SUCCESS
		mov		[esi].IoStatus.Information, 0

	.elseif [edi].Parameters.DeviceIoControl.IoControlCode == IOCTL_DISK_VERIFY
		mov		eax, [esi].AssociatedIrp.SystemBuffer
		assume eax:ptr VERIFY_INFORMATION
		mov		eax, [eax]._Length
		assume eax:ptr nothing
		mov		[esi].IoStatus.Information, eax
		mov		status, STATUS_SUCCESS
		assume edi:ptr IO_STACK_LOCATION

	.elseif ([edi].Parameters.DeviceIoControl.IoControlCode == IOCTL_DISK_CHECK_VERIFY) || ([edi].Parameters.DeviceIoControl.IoControlCode == IOCTL_STORAGE_CHECK_VERIFY) ||\
	  ([edi].Parameters.DeviceIoControl.IoControlCode == IOCTL_STORAGE_CHECK_VERIFY2)
		mov		[esi].IoStatus.Information, 0
		mov		status, STATUS_SUCCESS

	.elseif [edi].Parameters.DeviceIoControl.IoControlCode == IOCTL_DISK_IS_WRITABLE
		mov		status, STATUS_SUCCESS
		mov		[esi].IoStatus.Information, 0

	.elseif ([edi].Parameters.DeviceIoControl.IoControlCode == IOCTL_DISK_MEDIA_REMOVAL) || ([edi].Parameters.DeviceIoControl.IoControlCode == IOCTL_STORAGE_MEDIA_REMOVAL)
		mov		[esi].IoStatus.Information, 0
		mov		status, STATUS_SUCCESS

	.elseif [edi].Parameters.DeviceIoControl.IoControlCode == IOCTL_DISK_GET_PARTITION_INFO_EX
		.if [edi].Parameters.DeviceIoControl.OutputBufferLength >= sizeof PARTITION_INFORMATION_EX
		
			mov		edi, [esi].AssociatedIrp.SystemBuffer
			assume edi:ptr PARTITION_INFORMATION_EX
			mov		[edi].PartitionStyle, PARTITION_STYLE_MBR
			mov		eax, CreateParams.HiddenSectors
			
			mov		[edi].MBR.HiddenSectors, eax
			mul		CreateParams.BytesPerSector
			
			mov		[edi].StartingOffset.HighPart, 0
			mov		[edi].StartingOffset.LowPart, eax
			
			mov		[edi].PartitionLength.HighPart, 0
			mov		ecx, CreateParams.nDiskSize.LowPart
			sub		ecx, eax
			
			mov		[edi].PartitionLength.LowPart, ecx
			
			mov		[edi].PartitionNumber, 1
			
			mov		byte ptr [edi].RewritePartition, FALSE
			
			mov		byte ptr [edi].MBR.PartitionType, PARTITION_ENTRY_UNUSED
			
			mov		byte ptr [edi].MBR.BootIndicator, FALSE
			
			mov		byte ptr [edi].MBR.RecognizedPartition, FALSE
			
			mov		status, STATUS_SUCCESS
			mov		[esi].IoStatus.Information, sizeof(PARTITION_INFORMATION_EX)
			assume edi:ptr IO_STACK_LOCATION
		.else
			mov		status, STATUS_BUFFER_TOO_SMALL
			mov		[esi].IoStatus.Information, 0
		.endif
	.else
		mov		status, STATUS_INVALID_DEVICE_REQUEST
		mov		[esi].IoStatus.Information, 0
	.endif

	assume edi:nothing
ACDC_exit:
	mov		eax, status
	mov		[esi].IoStatus.Status, eax
	fastcall IofCompleteRequest, esi, IO_NO_INCREMENT

	mov		eax, status
	assume esi:nothing
	
	ret
AC_DispatchControl endp