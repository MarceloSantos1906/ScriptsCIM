import ctypes

def is_num_lock():
    return ctypes.windll.user32.GetKeyState(0x90) & 1

def get_capslock_state():
    hllDll = ctypes.WinDLL ("User32.dll")
    VK_CAPITAL = 0x14
    return hllDll.GetKeyState(VK_CAPITAL)

def get_numlock_state():
    hllDll = ctypes.WinDLL ("User32.dll")
    VK_NUMLOCK = 0x90
    return hllDll.GetKeyState(VK_NUMLOCK)

def get_scrolllock_state():
    hllDll = ctypes.WinDLL ("User32.dll")
    VK_SCROLL = 0x91
    return hllDll.GetKeyState(VK_SCROLL)
