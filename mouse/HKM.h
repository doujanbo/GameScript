#ifndef _HKM_
#define _HKM_

#define HKM_FAIL		((DWORD)-1) //ʧ��

#define HDT_ALL		0 //ȫ��ģʽ
#define HDT_KM		1 //����ģʽ
#define HDT_KEY		2 //����ģʽ
#define HDT_MOUSE	3 //���ģʽ

#define HMT_REL		1 //����������
#define HMT_ABS		2 //�����������
#define HMT_RA		3 //����������+�����������

#define HDEMI_POS		0x01 //��Ļ����
#define HDEMI_MOVE		0x02 //����ƶ���ʽ
#define HDEMI_WHEEL		0x03 //�����ֹ�����ʽ
#define HDEMI_STR		0x04 //����ַ�������
#define HDEMI_SYNC		0x05 //���̡����ͬ����ʽ
#define HDEMI_RND		0x06 //��ʱ�����ʽ
#define HDEMI_FUN		0x20 //�����Ĳ������߷���ֵ���ַ�������

#define HDEPM_PHYSICAL	0x00 //��������
#define HDEPM_VIRTUAL	0x01 //��������
#define HDEMM_DEFAULT	0x00 //Ĭ������ƶ���ʽ
#define HDEMM_QUICK		0x01 //�������ƶ�
#define HDEMM_ABS		0x02 //�������ƶ�
#define HDEWM_SIM		0x00 //�����ַ������
#define HDEWM_QUICK		0x01 //�����ֿ��ٹ���
#define HDESM_ANSI		0x00 //����ַ���ANSI���뷽ʽ
#define HDESM_UNICODE	0x01 //����ַ���UNICODE���뷽ʽ
#define HDESM_ANSI2		0x02 //����ַ���ANSI����2��ʽ
#define HDESM_UNICODE2	0x03 //����ַ���UNICODE����2��ʽ
#define HDEAM_SYNC		0x00 //����������ͬ��
#define HDEAM_KAS		0x01 //���̲����첽
#define HDEAM_MAS		0x02 //�������첽
#define HDERM_UFT		0x00 //��ʱ������ȷֲ�
#define HDERM_LITTLE	0x01 //��ʱ���ƫС�ֲ�
#define HDEFM_UNICODE	0x00 //�����Ĳ������߷���ֵ���ַ���ʹ��UNICODE����
#define HDEFM_ANSI		0x01 //�����Ĳ������߷���ֵ���ַ���ʹ��ANSI����

#define KBD_NUM_LOCK	0 //NumLockָʾ��
#define KBD_CAPS_LOCK	1 //CapsLockָʾ��
#define KBD_SCROLL_LOCK	2 //ScrollLockָʾ��

#define CKM_CHS			0x01 //��鰴�������ذ������ĵ���Ϣ
#define CKM_MOUSE		0x02 //��鰴�������ذ��������µ�������Ϣ
#define CKM_ANSI		0x04 //��鰴�������ص��ַ���ʹ��ANSI����

#ifdef __cplusplus
extern "C" {
#endif

DECLSPEC_IMPORT DWORD WINAPI HKMGetVersion(void);
DECLSPEC_IMPORT BOOL WINAPI HKMDisableDpi(void);
DECLSPEC_IMPORT int WINAPI HKMIsDisableDpi(void);
DECLSPEC_IMPORT DWORD WINAPI HKMSearchDevice(DWORD dwVid, DWORD dwPid, DWORD dwDeviceType);
DECLSPEC_IMPORT DWORD WINAPI HKMSearchDevice2(DWORD dwVid, DWORD dwPid, DWORD dwSN, DWORD dwDeviceType);
DECLSPEC_IMPORT LPVOID WINAPI HKMOpen(DWORD dwDeviceId);
DECLSPEC_IMPORT LPVOID WINAPI HKMOpen2(DWORD dwDeviceId1,DWORD dwDeviceId2);
DECLSPEC_IMPORT BOOL WINAPI HKMIsOpen(LPVOID lpHKMData,DWORD dwFlags);
DECLSPEC_IMPORT BOOL WINAPI HKMClose(LPVOID lpHKMData);
DECLSPEC_IMPORT DWORD WINAPI HKMGetSerialNumber(LPVOID lpHKMData,BOOL bMouse);
DECLSPEC_IMPORT DWORD WINAPI HKMGetMouseMode(LPVOID lpHKMData);
DECLSPEC_IMPORT BOOL WINAPI HKMIsKeyBusy(LPVOID lpHKMData);
DECLSPEC_IMPORT BOOL WINAPI HKMIsMouseBusy(LPVOID lpHKMData);
DECLSPEC_IMPORT BOOL WINAPI HKMGetKeyboardLEDState(LPVOID lpHKMData,DWORD dwIndex);
DECLSPEC_IMPORT BOOL WINAPI HKMGetCursorPos(LPVOID lpHKMData,LPLONG lpX,LPLONG lpY);
DECLSPEC_IMPORT DWORD WINAPI HKMGetCursorPos2(LPVOID lpHKMData);
DECLSPEC_IMPORT BOOL WINAPI HKMSetMode(LPVOID lpHKMData,DWORD dwIndex,DWORD dwMode);
DECLSPEC_IMPORT BOOL WINAPI HKMSetKeyInterval(LPVOID lpHKMData,DWORD dwMinTime,DWORD dwMaxTime);
DECLSPEC_IMPORT BOOL WINAPI HKMSetMouseInterval(LPVOID lpHKMData,DWORD dwMinTime,DWORD dwMaxTime);
DECLSPEC_IMPORT BOOL WINAPI HKMSetAbsMouseScrnRes(LPVOID lpHKMData,long lWidth,long lHeight);
DECLSPEC_IMPORT BOOL WINAPI HKMSetMousePosPrecision(LPVOID lpHKMData,DWORD dwPrecision);
DECLSPEC_IMPORT BOOL WINAPI HKMSetMouseMoveTimeout(LPVOID lpHKMData,DWORD dwTimeout);
DECLSPEC_IMPORT BOOL WINAPI HKMSetMousePosMaxOffset(LPVOID lpHKMData,DWORD dwOffset);
DECLSPEC_IMPORT BOOL WINAPI HKMKeyPress(LPVOID lpHKMData,LPCWSTR lpKeyName);
DECLSPEC_IMPORT BOOL WINAPI HKMKeyDown(LPVOID lpHKMData,LPCWSTR lpKeyName);
DECLSPEC_IMPORT BOOL WINAPI HKMKeyUp(LPVOID lpHKMData,LPCWSTR lpKeyName);
DECLSPEC_IMPORT BOOL WINAPI HKMMoveTo(LPVOID lpHKMData,long nX,long nY);
DECLSPEC_IMPORT BOOL WINAPI HKMMoveR(LPVOID lpHKMData,long nX,long nY);
DECLSPEC_IMPORT BOOL WINAPI HKMMoveR2(LPVOID lpHKMData,long nX,long nY);
DECLSPEC_IMPORT BOOL WINAPI HKMMoveRP(LPVOID lpHKMData,long nX,long nY);
DECLSPEC_IMPORT BOOL WINAPI HKMLeftClick(LPVOID lpHKMData);
DECLSPEC_IMPORT BOOL WINAPI HKMRightClick(LPVOID lpHKMData);
DECLSPEC_IMPORT BOOL WINAPI HKMMiddleClick(LPVOID lpHKMData);
DECLSPEC_IMPORT BOOL WINAPI HKMLeftDoubleClick(LPVOID lpHKMData);
DECLSPEC_IMPORT BOOL WINAPI HKMRightDoubleClick(LPVOID lpHKMData);
DECLSPEC_IMPORT BOOL WINAPI HKMMiddleDoubleClick(LPVOID lpHKMData);
DECLSPEC_IMPORT BOOL WINAPI HKMLeftDown(LPVOID lpHKMData);
DECLSPEC_IMPORT BOOL WINAPI HKMRightDown(LPVOID lpHKMData);
DECLSPEC_IMPORT BOOL WINAPI HKMMiddleDown(LPVOID lpHKMData);
DECLSPEC_IMPORT BOOL WINAPI HKMLeftUp(LPVOID lpHKMData);
DECLSPEC_IMPORT BOOL WINAPI HKMRightUp(LPVOID lpHKMData);
DECLSPEC_IMPORT BOOL WINAPI HKMMiddleUp(LPVOID lpHKMData);
DECLSPEC_IMPORT BOOL WINAPI HKMMouseWheel(LPVOID lpHKMData,long nCount);
DECLSPEC_IMPORT BOOL WINAPI HKMMouseWheelP(LPVOID lpHKMData,long nCount);
DECLSPEC_IMPORT BOOL WINAPI HKMReleaseKeyboard(LPVOID lpHKMData);
DECLSPEC_IMPORT BOOL WINAPI HKMReleaseMouse(LPVOID lpHKMData);
DECLSPEC_IMPORT BOOL WINAPI HKMOutputString(LPVOID lpHKMData,LPCWSTR lpString);
DECLSPEC_IMPORT BOOL WINAPI HKMDelayRnd(LPVOID lpHKMData,DWORD dwMinTime,DWORD dwMaxTime);
DECLSPEC_IMPORT LPWSTR WINAPI HKMCheckPressedKeys(DWORD dwFlags,LPDWORD lpLength);
DECLSPEC_IMPORT BOOL WINAPI HKMFreeData(LPVOID lpData);
DECLSPEC_IMPORT BOOL WINAPI HKMVerifyUserData(LPVOID lpHKMData,LPCWSTR lpString,BOOL bMouse);
DECLSPEC_IMPORT BOOL WINAPI HKMIsOSMouseAccelerateEnabled(void);
DECLSPEC_IMPORT BOOL WINAPI HKMEnableOSMouseAccelerate(BOOL bEnable,BOOL bSave);
DECLSPEC_IMPORT int WINAPI HKMGetOSMouseSpeed(void);
DECLSPEC_IMPORT BOOL WINAPI HKMSetOSMouseSpeed(int nSpeed,BOOL bSave);

#ifdef __cplusplus
}
#endif

#endif //_HKM_
