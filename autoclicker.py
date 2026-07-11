import threading
import time
import keyboard
import pyautogui

pyautogui.FAILSAFE = False

active = False
running = True

def toggle():
    global active
    active = not active
    estado = "ACTIVADO" if active else "DESACTIVADO"
    print(f"[F2] Auto clicker {estado}")

def click_loop():
    while running:
        if active:
            pyautogui.click()
        time.sleep(0.001)  # ~1000 clicks/seg

keyboard.add_hotkey("f2", toggle)

print("Auto Clicker listo")
print("F2  → activar / desactivar")
print("ESC → salir\n")

click_thread = threading.Thread(target=click_loop, daemon=True)
click_thread.start()

keyboard.wait("esc")
running = False
print("Cerrando...")
