"""
Capture Wireshark screenshots v4 - FIXED:
- Properly expands packet detail tree (End->Right, navigate to IP->Right, ICMP->Right)
- No Ctrl+Shift+E (that was Enabled Protocols dialog)
- Suppress save dialog
- Force topmost
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
EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.wintypes.HWND, ctypes.wintypes.LPARAM)
SW_SHOWMAXIMIZED = 3
SWP_NOMOVE = 0x0002
SWP_NOSIZE = 0x0001
HWND_TOPMOST = -1
HWND_NOTOPMOST = -2

def find_ws():
    result = []
    def cb(hwnd, _):
        if user32.IsWindowVisible(hwnd):
            buf = ctypes.create_unicode_buffer(512)
            user32.GetWindowTextW(hwnd, buf, 512)
            if 'wireshark' in buf.value.lower():
                result.append(hwnd)
        return True
    user32.EnumWindows(EnumWindowsProc(cb), 0)
    return result[0] if result else None

def force_front(hwnd):
    user32.SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE)
    time.sleep(0.15)
    user32.ShowWindow(hwnd, SW_SHOWMAXIMIZED)
    time.sleep(0.2)
    user32.SetForegroundWindow(hwnd)
    time.sleep(0.2)
    user32.BringWindowToTop(hwnd)
    time.sleep(0.3)
    user32.SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE)
    time.sleep(0.15)
    user32.ShowWindow(hwnd, SW_SHOWMAXIMIZED)
    time.sleep(0.5)

def expand_tree(hwnd):
    """Expand IP and ICMP protocol headers in the packet detail pane."""
    user32.SetForegroundWindow(hwnd)
    time.sleep(0.3)
    sw = user32.GetSystemMetrics(0)
    sh = user32.GetSystemMetrics(1)

    # Click in the packet detail pane (roughly 50% down the screen, left side)
    pyautogui.click(sw // 4, int(sh * 0.50))
    time.sleep(0.4)

    # Strategy: go to end (last top-level item = ICMP or last protocol), expand it
    # Then go to top, skip Frame & Ethernet, expand IP
    # All top-level items when collapsed: Frame, Ethernet, IP, ICMP (4 items)

    # Go to last item (ICMP) and expand
    pyautogui.press('end')
    time.sleep(0.2)
    pyautogui.press('right')  # expand ICMP
    time.sleep(0.2)

    # Go to top
    pyautogui.press('home')
    time.sleep(0.2)
    # Frame is selected. Skip it.
    pyautogui.press('down')  # -> Ethernet
    time.sleep(0.1)
    pyautogui.press('down')  # -> IP
    time.sleep(0.1)
    pyautogui.press('right')  # expand IP
    time.sleep(0.2)

    # Scroll to show everything - press Home to go to top of tree
    pyautogui.press('home')
    time.sleep(0.3)

def grab_label(name, labels):
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

def capture(trace, filt, pkt, name, labels, do_expand=True):
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
        hwnd = find_ws()
        if hwnd:
            break
    if not hwnd:
        print(f"  FAILED")
        proc.kill()
        return None
    time.sleep(5)
    force_front(hwnd)
    time.sleep(1)
    if do_expand:
        expand_tree(hwnd)
        time.sleep(0.5)
        force_front(hwnd)
        time.sleep(0.5)
    path = grab_label(name, labels)
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
    "ICMP Q9: Type=11 (TTL-exceeded) from router 192.168.68.1",
    "Extra fields: encapsulated original IP hdr + 8B of triggering pkt"
])

print("\n[ 5/11] ICMP Q9: Last packets + ping reply for comparison")
capture(PING, "icmp", 580, "icmp_q9_echo_reply", [
    "ICMP Q9: Echo Reply (Type=0) - what last 3 tracert pkts look like",
    "vs TTL-exceeded (Type=11): no encapsulated headers, from destination"
])

print("\n[ 6/11] IP Q3: 56-byte ping")
capture(COMBINED, "icmp", 61, "ip_q3_56byte", [
    "IP Q3: Header=20B (IHL=5) | Total=84B | Payload=84-20=64B",
    "Protocol=ICMP(1) | Not fragmented (Flags=0x00)"
])

print("\n[ 7/11] IP Q5+Q7: Tracert requests showing TTL+ID changes")
capture(TRACERT, "icmp.type==8", None, "ip_q5_q7_tracert_pattern", [
    "IP Q5: Fields that CHANGE: TTL (1->2->3...), Identification, Checksum",
    "IP Q7: ID increments: 0xda8a -> 0xda8b -> ... (sequential)"
])

print("\n[ 8/11] IP Q9: First-hop replies")
capture(TRACERT, "icmp.type==11 && ip.src==192.168.68.1", 116, "ip_q9_first_hop", [
    "IP Q9: First-hop (192.168.68.1) replies | TTL constant=64",
    "ID changes: 0x1f50, 0x1f54, 0x1f57 (router's own counter)"
])

print("\n[ 9/11] IP Q11: First fragment")
capture(COMBINED, "ip.id==0xdab6", 1337, "ip_q11_first_fragment", [
    "IP Q11: 1st fragment | Len=1500 | MF=1 | Offset=0",
    "ID=0xdab6 | First: Offset=0 + MF=1"
])

print("\n[10/11] IP Q11: Second fragment")
capture(COMBINED, "ip.id==0xdab6", 1338, "ip_q11_second_fragment", [
    "IP Q11: 2nd/last fragment | Len=548 | MF=0 | Offset=1480",
    "Same ID=0xdab6 | Reassembled=2008B"
])

print("\n[11/11] IP Q15: 3500B fragments")
capture(COMBINED, "ip.id==0xdabb", 1513, "ip_q15_3500_fragments", [
    "IP Q15: 3 fragments | ID=0xdabb",
    "Len:1500,1500,568 | MF:1,1,0 | Offset:0,1480,2960"
])

print("\n" + "="*50)
print(" ALL 11 DONE!")
print("="*50)
