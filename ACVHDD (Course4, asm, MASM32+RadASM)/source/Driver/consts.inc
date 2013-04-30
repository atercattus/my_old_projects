.386
.model flat, stdcall
option casemap:none

include \masm32\include\w2k\ntstatus.inc
include \masm32\include\w2k\ntddk.inc
include \masm32\include\w2k\ntoskrnl.inc
include \masm32\include\w2k\ntdddisk.inc

include \masm32\Macros\Strings.mac
include \masm32\Macros\Mutex.mac
include proto.inc
include seh0.inc

includelib \masm32\lib\w2k\ntoskrnl.lib

RAMDISK_MEDIA_TYPE      equ 0F8h
DIR_ENTRIES_PER_SECTOR  equ 112h

BYTES_PER_SECTOR		equ 512
SECTORS_PER_TRACK		equ 32
TRACKS_PER_CYLINDER		equ 2
HIDDEN_SECTORS			equ 1

TokenPrimary			equ	1
TokenImpersonation		equ	2

MOUNTDEVCONTROLTYPE		equ	"M"
FTTYPE					equ	"f"
FILE_DEVICE_FILE_SYSTEM	equ 9

START_CRYPT				equ	1

IOCTL_MOUNTDEV_QUERY_DEVICE_NAME	equ CTL_CODE(MOUNTDEVCONTROLTYPE, 2, METHOD_BUFFERED, FILE_ANY_ACCESS)
FT_BALANCED_READ_MODE				equ	CTL_CODE(FTTYPE, 6, METHOD_NEITHER,  FILE_ANY_ACCESS)
IOCTL_STORAGE_GET_HOTPLUG_INFO		equ CTL_CODE(IOCTL_STORAGE_BASE, 0305h, METHOD_BUFFERED, FILE_ANY_ACCESS)
IOCTL_DISK_GET_PARTITION_INFO_EX	equ CTL_CODE(IOCTL_DISK_BASE, 12h, METHOD_BUFFERED, FILE_ANY_ACCESS)
IOCTL_UNMOUNT_DISK					equ CTL_CODE(FILE_DEVICE_UNKNOWN, 1, METHOD_BUFFERED, FILE_ANY_ACCESS)
IOCTL_MOUNT_DISK					equ CTL_CODE(FILE_DEVICE_UNKNOWN, 2, METHOD_BUFFERED, FILE_ANY_ACCESS)
IOCTL_SET_KEY_DISK					equ CTL_CODE(FILE_DEVICE_UNKNOWN, 3, METHOD_BUFFERED, FILE_ANY_ACCESS)

IOCTL_DISK_GET_LENGTH_INFO			equ CTL_CODE(IOCTL_DISK_BASE, 17h, METHOD_BUFFERED, FILE_READ_ACCESS)

FSCTL_DISMOUNT_VOLUME				equ CTL_CODE(FILE_DEVICE_FILE_SYSTEM, 8, METHOD_BUFFERED, FILE_ANY_ACCESS)
FSCTL_LOCK_VOLUME					equ CTL_CODE(FILE_DEVICE_FILE_SYSTEM, 6, METHOD_BUFFERED, FILE_ANY_ACCESS)
FSCTL_UNLOCK_VOLUME					equ CTL_CODE(FILE_DEVICE_FILE_SYSTEM, 7, METHOD_BUFFERED, FILE_ANY_ACCESS)

FSCTL_IS_VOLUME_MOUNTED				equ CTL_CODE(FILE_DEVICE_FILE_SYSTEM, 10, METHOD_BUFFERED, FILE_ANY_ACCESS)

PARTITION_STYLE_MBR		equ	0

GET_LENGTH_INFORMATION STRUCT
	_Length		LARGE_INTEGER <>
GET_LENGTH_INFORMATION ENDS
PGET_LENGTH_INFORMATION typedef ptr GET_LENGTH_INFORMATION

GUID	struct
	Data1		dd	?
	Data2		dw	?
	Data3		dw	?
	Data4		db	8 dup(?)
GUID	ends
	;
PARTITION_INFORMATION_MBR	struct			; size = 7h
	PartitionType				UCHAR		?
	BootIndicator				BOOLEAN		?
	RecognizedPartition			BOOLEAN		?
	align 2
	HiddenSectors				ULONG		?
PARTITION_INFORMATION_MBR	ends
	;
PARTITION_INFORMATION_GPT	struct			; size = 70h
	PartitionType				GUID 			<>
	PartitionId					GUID			<>
	Attributes					LARGE_INTEGER	<>
	_Name						WORD			36 dup(?)
PARTITION_INFORMATION_GPT	ends
	;
PARTITION_INFORMATION_EX	struct
	PartitionStyle		DWORD			?
	align 8
	StartingOffset		LARGE_INTEGER	<>
	PartitionLength		LARGE_INTEGER	<>
	PartitionNumber		ULONG			?
	RewritePartition	BOOLEAN			?
	align 4
	union
		MBR					PARTITION_INFORMATION_MBR	<>
		GPT					PARTITION_INFORMATION_GPT	<>
	ends
PARTITION_INFORMATION_EX	ends
PPARTITION_INFORMATION_EX	typedef ptr PARTITION_INFORMATION_EX
	;
CREATE_PARAMS	struct
	nDiskSize				LARGE_INTEGER   <>
	HiddenSectors			DWORD	?
	Cylinders               LARGE_INTEGER   <>
	MediaType               DWORD   ?
	TracksPerCylinder       DWORD   ?
	SectorsPerTrack         DWORD   ?
	BytesPerSector          DWORD   ?
CREATE_PARAMS	ends

.const
	CCOUNTED_UNICODE_STRING	"\\Device\\ACVirtDrive", usDeviceName, 4

.data?
	pSyncStart				PVOID			?
	nOpenCount				DWORD			?
	hDiskFile				HANDLE			?
	pMemory					PVOID			?
	BufferLength			DWORD			?
;	ReadOnly				BOOL			?
	CreateParams			CREATE_PARAMS	<>
	kDiskMutex				KMUTEX			<>