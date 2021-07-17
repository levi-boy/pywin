import win32gui
import win32process
import win32api
import win32com
import pygetwindow


def find_window_and_set_focus(title):
    try:
        hwnd = win32gui.FindWindow(None, title)
        remote_thread, _ = win32process.GetWindowThreadProcessId(hwnd)
        win32process.AttachThreadInput(win32api.GetCurrentThreadId(), remote_thread, True)
        win32gui.SetFocus(hwnd)
        shell = win32com.client.Dispatch("WScript.Shell")
        shell.SendKeys('%')
        win32gui.SetForegroundWindow(hwnd)
    except Exception:
        pass


def get_window_panel(desired_panel):
    atomic_panel = [current_panel for current_panel in pygetwindow.getAllTitles()
        if current_panel.startswith(desired_panel)]
    return {
        'width': pygetwindow.getWindowsWithTitle(atomic_panel[0])[0].width,
        'height': pygetwindow.getWindowsWithTitle(atomic_panel[0])[0].height,
        'x1': pygetwindow.getWindowsWithTitle(atomic_panel[0])[0].left,
        'y1': pygetwindow.getWindowsWithTitle(atomic_panel[0])[0].top
    }, pygetwindow.getWindowsWithTitle(atomic_panel[0])[0]


def restore_and_resize_panel(title):
    _, panel = get_window_panel(title)
    panel.restore()
    panel.resizeTo(1110, 730)