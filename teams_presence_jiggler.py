import argparse
import ctypes
import signal
import sys
import time


class POINT(ctypes.Structure):
    _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]


def get_cursor_pos():
    pt = POINT()
    ctypes.windll.user32.GetCursorPos(ctypes.byref(pt))
    return pt.x, pt.y


def set_cursor_pos(x, y):
    ctypes.windll.user32.SetCursorPos(x, y)


def press_key():
    """Press and release Shift key to simulate user input"""
    # VK_SHIFT = 0x10
    ctypes.windll.user32.keybd_event(0x10, 0, 0, 0)  # Key down
    time.sleep(0.05)
    ctypes.windll.user32.keybd_event(0x10, 0, 2, 0)  # Key up


def jiggler(interval: float, distance: int):
    start_time = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
    print(f"Teams presence jiggler started at {start_time}. Press Ctrl+C to stop.")
    print(f"Settings: interval={interval:.0f}s, mouse_distance={distance}px")
    tick = 0
    try:
        while True:
            tick += 1
            # Move mouse
            x, y = get_cursor_pos()
            set_cursor_pos(x + distance, y)
            time.sleep(0.1)
            set_cursor_pos(x, y)
            
            # Simulate key press (Shift)
            press_key()
            time.sleep(0.1)
            
            current_time = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
            print(f"[{current_time}] action #{tick}: moved cursor ({distance}px) + key press. waiting {interval:.0f}s...")
            time.sleep(interval)
    except KeyboardInterrupt:
        print("\nStopped.")



def main():
    parser = argparse.ArgumentParser(
        description="Prevent Teams idle/away state by gently moving the mouse cursor."
    )
    parser.add_argument(
        "-i",
        "--interval",
        type=float,
        default=60.0,
        help="Seconds between cursor nudges (default: 60)",
    )
    parser.add_argument(
        "-d",
        "--distance",
        type=int,
        default=5,
        help="Cursor move distance in pixels (default: 5)",
    )
    args = parser.parse_args()

    if args.interval <= 0:
        parser.error("Interval must be greater than 0")

    jiggler(args.interval, args.distance)


if __name__ == "__main__":
    main()
