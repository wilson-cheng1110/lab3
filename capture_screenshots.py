"""
Capture Wireshark screenshots - fixed version.
No Ctrl+Shift+E (that opens Enabled Protocols dialog).
Instead: click on detail pane items and press Right Arrow to expand.
Also suppress the save dialog with -o gui.ask_unsaved:FALSE.
"""
import subprocess
import time
import os
import ctypes
import ctypes.wintypes
import pyautogui
from PIL import ImageGrab, ImageDraw, ImageFont

pyautogui.FAILSAFE = False

WIRESHARK = r"C:\Program Files\Wireshark\Wireshark.exe"
OUTDIR = r"C:\Users\User\Downloads\lab3\screenshots"
os.makedirs(OUTDIR, exist_ok=True)

user32 = ctypes.windll.user32
EnumWindows = user32.EnumWindows
EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.wintypes.HWND, ctypes.wintypes.LPARAM)
GetWindowText = user32.GetWindowTextW
IsWindowVisible = user32.IsWindowVisible
ShowWindow = user32.ShowWindow
SetForegroundWindow = user32.SetForegroundWindow
BringWindowToTop = user32.BringWindowToTop
SW_SHOWMAXIMIZED = 3
SWP_NOMOVE = 0x0002
SWP_NOSIZE = 0x0001
HWND_TOPMOST = -1
HWND_NOTOPMOST = -2

def find_ws_hwnd():
    result = []
    def callback(hwnd, lParam):
        if IsWindowVisible(hwnd):
            title = ctypes.create_unicode_buffer(512)
            GetWindowText(hwnd, title, 512)
            t = title.value.lower()
            if 'wireshark' in t:
                result.append(hwnd)
        return True
    EnumWindows(EnumWindowsProc(callback), 0)
    return result[0] if result else None

def force_front(hwnd):
    user32.SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE)
    time.sleep(0.2)
    ShowWindow(hwnd, SW_SHOWMAXIMIZED)
    time.sleep(0.3)
    SetForegroundWindow(hwnd)
    time.sleep(0.3)
    BringWindowToTop(hwnd)
    time.sleep(0.5)
    user32.SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE)
    time.sleep(0.2)
    ShowWindow(hwnd, SW_SHOWMAXIMIZED)
    time.sleep(0.5)

def expand_details(hwnd):
    """Click on packet detail pane and expand IP + ICMP headers."""
    SetForegroundWindow(hwnd)
    time.sleep(0.3)
    sw = user32.GetSystemMetrics(0)
    sh = user32.GetSystemMetrics(1)

    # The packet detail pane is roughly in the middle 40-70% of the screen height.
    # Click on the first item in the detail tree (Frame header) ~ 42% down
    detail_x = sw // 4
    detail_top = int(sh * 0.42)

    # Click on detail pane to focus it
    pyautogui.click(detail_x, detail_top)
    time.sleep(0.3)

    # Press Home to go to first item
    pyautogui.press('home')
    time.sleep(0.2)

    # Expand Frame (press right arrow)
    pyautogui.press('right')
    time.sleep(0.1)

    # Move down to Ethernet header
    pyautogui.press('down')
    time.sleep(0.1)
    # Expand Ethernet
    pyautogui.press('right')
    time.sleep(0.1)

    # Move down to IP header
    pyautogui.press('down')
    time.sleep(0.1)
    # Find IP - press down a few times and expand
    for _ in range(5):
        pyautogui.press('down')
        time.sleep(0.05)
    # Now try expanding - press right
    pyautogui.press('right')
    time.sleep(0.1)

    # Move down more to ICMP
    for _ in range(15):
        pyautogui.press('down')
        time.sleep(0.05)
    pyautogui.press('right')
    time.sleep(0.1)

    # Scroll up a bit to show the full details
    pyautogui.press('home')
    time.sleep(0.2)

def grab_and_label(name, labels):
    img = ImageGrab.grab()
    draw = ImageDraw.Draw(img)
    try:
        font = ImageFont.truetype("C:/Windows/Fonts/arialbd.ttf", 20)
    except:
        font = ImageFont.load_default()

    y = 32
    for label in labels:
        bbox = draw.textbbox((0, 0), label, font=font)
        tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
        draw.rectangle([8, y, tw + 24, y + th + 12], fill='red', outline='darkred', width=2)
        draw.text((14, y + 6), label, fill='white', font=font)
        y += th + 18

    path = os.path.join(OUTDIR, f"{name}.png")
    img.save(path)
    print(f"  -> {path}")
    return path

def capture(trace, filt, pkt, name, labels):
    os.system('taskkill /IM Wireshark.exe /F >nul 2>&1')
    time.sleep(2)

    cmd = [WIRESHARK, "-r", trace, "-o", "gui.ask_unsaved:FALSE"]
    if filt:
        cmd += ["-Y", filt]
    if pkt:
        cmd += ["-g", str(pkt)]

    proc = subprocess.Popen(cmd)
    print(f"  Loading...")

    hwnd = None
    for _ in range(60):
        time.sleep(0.25)
        hwnd = find_ws_hwnd()
        if hwnd:
            break

    if not hwnd:
        print(f"  FAILED: window not found")
        proc.kill()
        return None

    time.sleep(4)
    force_front(hwnd)
    time.sleep(1)

    # Expand packet details
    expand_details(hwnd)
    time.sleep(0.5)

    # Force front again
    force_front(hwnd)
    time.sleep(1)

    path = grab_and_label(name, labels)

    # Close - no save dialog thanks to gui.ask_unsaved:FALSE
    user32.PostMessageW(hwnd, 0x0010, 0, 0)
    time.sleep(2)
    try:
        proc.kill()
    except:
        pass
    time.sleep(1)
    return path

PING = r"C:\Users\User\Downloads\lab3\lab3_icmp_ping.pcapng"
TRACERT = r"C:\Users\User\Downloads\lab3\lab3_icmp_tracert.pcapng"
COMBINED = r"C:\Users\User\Downloads\lab3\lab3_ip_combined.pcapng"

print("="*50)
print(" Capturing 11 screenshots. Don't touch mouse!")
print("="*50)

print("\n[ 1/11] ICMP Q1+Q3: Ping request")
capture(PING, "icmp", 579, "icmp_q1_q3_ping_request", [
    "ICMP Q1+Q3: Src=192.168.68.75  Dst=143.89.209.9 (www.ust.hk)",
    "Type=8 (Echo Request), Code=0 | Checksum=2B, ID=2B, Seq=2B"
])

print("\n[ 2/11] ICMP Q3: Ping reply")
capture(PING, "icmp", 580, "icmp_q3_ping_reply", [
    "ICMP Q3: Echo Reply | Type=0, Code=0",
    "Checksum=0x090b(2B), ID=0x0007(2B), Seq=19530(2B)"
])

print("\n[ 3/11] ICMP Q7: Tracert echo")
capture(TRACERT, "icmp", 115, "icmp_q7_tracert_echo", [
    "ICMP Q7: Tracert Echo Request - TTL=1 (ping uses 128)",
    "Data=64B zeros (ping=32B 'abcdef...'), IP Total=92 (ping=60)"
])

print("\n[ 4/11] ICMP Q9: TTL-exceeded")
capture(TRACERT, "icmp", 116, "icmp_q9_ttl_exceeded", [
    "ICMP Q9: Type=11 (TTL-exceeded), Code=0 from 192.168.68.1",
    "Contains original IP header + ICMP header of triggering packet"
])

print("\n[ 5/11] ICMP Q9: Last 3 received")
capture(TRACERT, "icmp && ip.dst==192.168.68.75", 4652, "icmp_q9_last_packets", [
    "ICMP Q9: Last 3 = TTL-exceeded from hop 9 (213.144.183.207)",
    "Complete tracert: these would be Echo Reply (Type=0) from destination"
])

print("\n[ 6/11] IP Q3: 56-byte ping")
capture(COMBINED, "icmp", 61, "ip_q3_56byte", [
    "IP Q3: Header=20B (IHL=5) | Total=84B | Payload=84-20=64B",
    "Protocol=ICMP(1) | Flags=0x00 -> Not fragmented"
])

print("\n[ 7/11] IP Q5+Q7: ID pattern")
capture(PING, "icmp.type==8", None, "ip_q5_q7_id_pattern", [
    "IP Q5: Identification & Checksum change each packet",
    "IP Q7: ID increments +1: 0x0b71 -> 0x0b72 -> ... -> 0x0b7a"
])

print("\n[ 8/11] IP Q9: First-hop replies")
capture(TRACERT, "icmp.type==11 && ip.src==192.168.68.1", 116, "ip_q9_first_hop", [
    "IP Q9: First-hop (192.168.68.1) TTL-exceeded replies",
    "ID changes (0x1f50,0x1f54,0x1f57) | TTL constant=64"
])

print("\n[ 9/11] IP Q11: First fragment")
capture(COMBINED, "ip.id==0xdab6", 1337, "ip_q11_first_fragment", [
    "IP Q11: 1st fragment | Len=1500 | MF=1 | Offset=0",
    "ID=0xdab6 | First fragment: Offset=0 and MF=1"
])

print("\n[10/11] IP Q11: Second fragment")
capture(COMBINED, "ip.id==0xdab6", 1338, "ip_q11_second_fragment", [
    "IP Q11: 2nd/last fragment | Len=548 | MF=0 | Offset=1480",
    "Same ID=0xdab6 | Reassembled = 2008 bytes"
])

print("\n[11/11] IP Q15: 3500B fragments")
capture(COMBINED, "ip.id==0xdabb", 1513, "ip_q15_3500_fragments", [
    "IP Q15: 3 fragments | ID=0xdabb",
    "Len:1500,1500,568 | MF:1,1,0 | Offset:0,1480,2960"
])

print("\n" + "="*50)
print(" ALL 11 DONE!")
print("="*50)
