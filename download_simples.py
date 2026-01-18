import pyautogui
import time
import os
import sys
import pyperclip
import random
import math

# --- CONFIGURATION ---
CNPJ_LIST = ["11575133000102", "22190596000172"] 
CNPJ_LIST = ["76486455000120",
"77620631000138",
"78346343000108",
"79958856000124",
"80361405000194",
"82060922000159",
"00541179000194",
"02011681000119",
"02330236000111",
"02358019000130",
"02530770000171",
"02948471000151",
"03441246000197",
"03457226000104"
]
ASSETS_DIR = os.path.join(os.getcwd(), "assets")
pyautogui.FAILSAFE = True 

# --- HUMAN PHYSICS (FASTER) ---

def human_sleep(min_seconds=0.1, max_seconds=0.4):
    """Very fast 'thinking' time."""
    time.sleep(random.uniform(min_seconds, max_seconds))

def get_point_on_curve(p0, p1, p2, t):
    x = (1-t)**2 * p0[0] + 2*(1-t)*t * p1[0] + t**2 * p2[0]
    y = (1-t)**2 * p0[1] + 2*(1-t)*t * p1[1] + t**2 * p2[1]
    return (x, y)

def human_move_to(target_x, target_y):
    """Moves mouse in a curve efficiently."""
    start_x, start_y = pyautogui.position()
    dist = math.hypot(target_x - start_x, target_y - start_y)
    
    offset = random.randint(int(dist * 0.1), int(dist * 0.3))
    direction = 1 if random.random() > 0.5 else -1
    mid_x = (start_x + target_x) / 2
    mid_y = (start_y + target_y) / 2
    control_x = mid_x + random.randint(-offset, offset)
    control_y = mid_y + (offset * direction)

    steps = random.randint(8, 15) 
    
    for i in range(steps):
        t = i / steps
        next_x, next_y = get_point_on_curve((start_x, start_y), (control_x, control_y), (target_x, target_y), t)
        pyautogui.moveTo(next_x, next_y, duration=0.002)

    pyautogui.moveTo(target_x, target_y, duration=0.05)

def get_jitter_point(location):
    left, top, width, height = location
    padding_x = int(width * 0.25)
    padding_y = int(height * 0.25)
    rand_x = random.randint(left + padding_x, left + width - padding_x)
    rand_y = random.randint(top + padding_y, top + height - padding_y)
    return rand_x, rand_y

def find_and_act(image_name, action="click", text=None, retries=10, offset_x=0, offset_y=0):
    img_path = os.path.join(ASSETS_DIR, image_name)
    if not os.path.exists(img_path):
        print(f"❌ MISSING ASSET: {image_name}")
        return False

    print(f"🔎 Scanning '{image_name}'...")
    
    for i in range(retries):
        try:
            location = pyautogui.locateOnScreen(img_path, confidence=0.8, grayscale=True)
            if location:
                target_x, target_y = get_jitter_point(location)
                target_x += offset_x
                target_y += offset_y
                
                human_move_to(target_x, target_y)
                
                if action == "click":
                    pyautogui.click()
                
                elif action == "type":
                    pyautogui.click()
                    time.sleep(0.1)
                    pyautogui.hotkey('ctrl', 'a')
                    pyautogui.press('backspace')
                    for char in text:
                        pyautogui.write(char)
                        time.sleep(random.uniform(0.01, 0.05))
                
                elif action == "triple_click":
                    pyautogui.click(clicks=3, interval=0.05)
                
                return True
            time.sleep(0.2)
        except Exception:
            pass
    return False

def clean_filename(text):
    invalid = '<>:"/\\|?*'
    for char in invalid:
        text = text.replace(char, '')
    text = text.replace('Nome Empresarial: ', '')
    return text.strip()

def perform_scroll_strategy():
    print("   📜 Scrolling...")
    screen_w, screen_h = pyautogui.size()
    
    # Inner Frame
    human_move_to(screen_w / 2, screen_h / 2)
    pyautogui.click()
    pyautogui.scroll(-1000)
    
    # Outer Frame
    human_move_to(screen_w - 20, screen_h / 2)
    pyautogui.scroll(-1500)
    
    # End Key
    pyautogui.click(screen_w / 2, screen_h / 2) 
    pyautogui.press('end')

def process_company(cnpj):
    print(f"\n🔵 PROCESSING CNPJ: {cnpj}")
    
    # REFRESH (Start Clean)
    pyautogui.press('f5')
    time.sleep(random.uniform(4.0, 6.0)) # Wait for reload

    # 1. SEARCH (Type CNPJ)
    if not find_and_act("input.png", action="type", text=cnpj):
        print("❌ Input missing.")
        return False

    # --- CHANGE: PRESS ENTER INSTEAD OF CLICKING BUTTON ---
    print("   Typing Enter...")
    human_sleep(0.2, 0.4) # Brief human pause after typing
    pyautogui.press('enter')
    # ------------------------------------------------------

    # 2. NAME EXTRACT
    time.sleep(2.0) # Wait for results
    if find_and_act("label_nome.png", action="triple_click", retries=15):
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(0.1)
        try:
            company_name = clean_filename(pyperclip.paste())
            print(f"   ✅ Name: {company_name}")
        except:
            company_name = "Company_Unknown"
    else:
        print("❌ Name label missing (Possible Captcha Block).")
        return False

    # 3. EXPAND INFO
    if find_and_act("mais_info.png", action="click", retries=10):
        time.sleep(1.0)
        perform_scroll_strategy()
        
        # 4. DOWNLOAD
        if find_and_act("btn_pdf.png", action="click", retries=10):
            print("   Saving PDF...")
            time.sleep(2.0) # Wait for dialog
            
            pyautogui.write(company_name, interval=0.01)
            pyautogui.press('enter')
            time.sleep(3.0) # Wait for download
            
            # 5. RESET (Voltar -> Scroll Up)
            print("   🔙 Resetting view...")
            if find_and_act("voltar.png", action="click", retries=10):
                time.sleep(1.5)
                pyautogui.press('home')
                
                # Reset Outer Scrollbar
                screen_w, screen_h = pyautogui.size()
                pyautogui.moveTo(screen_w - 20, screen_h / 2)
                pyautogui.scroll(2000)
                time.sleep(0.5)
                return True
            else:
                print("❌ Voltar missing.")
                return False
        else:
            print("❌ PDF Button missing.")
            return False
    else:
        print("❌ Expand missing.")
        return False

def main():
    print("="*50)
    print("   ROBUST AUTOMATION (ENTER KEY EDITION)")
    print("   Switch to Chrome NOW.")
    print("="*50)

    for k in range(3, 0, -1):
        print(f"{k}...")
        time.sleep(1)

    for i, cnpj in enumerate(CNPJ_LIST):
        print(f"--- Item {i+1}/{len(CNPJ_LIST)} ---")
        
        success = False
        attempt = 0
        max_retries = 3
        
        while not success and attempt < max_retries:
            if attempt > 0:
                print(f"⚠️ Error detected. Retrying ({attempt}/{max_retries})...")
                wait_time = random.uniform(10.0, 15.0)
                print(f"   Thinking for {wait_time:.1f}s...")
                time.sleep(wait_time)
            
            success = process_company(cnpj)
            attempt += 1
        
        if not success:
            print(f"⛔ CRITICAL: Failed to process {cnpj} after {max_retries} attempts. Skipping.")
        else:
            time.sleep(random.uniform(2.0, 4.0))

    print("\n🎉 ALL DONE!")

if __name__ == "__main__":
    main()