import os
import sys
import tkinter as tk
from tkinter import messagebox, simpledialog
import webbrowser
from PIL import Image, ImageTk
import pyttsx3
import speech_recognition as sr
import ctypes as ct
from tkinter import ttk
from screeninfo import get_monitors
import socket
import tkinter.filedialog as fd
import time
import threading
import pygame
import pyautogui
import string

# Global variables
voice_interaction_enabled = False
fullscreen = False
menu_visible = False
stop_listening = None
sound_muted = False
menu_animation_id = None
typing_active = False
last_typed_position = None

# Initialize speech recognition with error handling
recognizer = sr.Recognizer()
try:
    microphone = sr.Microphone()
    print("Microphone initialized successfully")
except OSError as e:
    print(f"Microphone initialization error: {e}")

# Indicator settings
indicator_colors = {
    'active': 'white',
    'inactive': 'black',
    'outline': 'black'
}

# Initialize pyttsx3 engine
engine = pyttsx3.init()
engine.setProperty('rate', 150)

# Initialize pygame mixer
pygame.mixer.init()

def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(sys.argv[0])
    return os.path.join(base_path, relative_path)

# Default sound paths
bing_sound_path = resource_path('Button.mp3')
menu_sound_path = resource_path('Click.mp3')
Hover_sound_path = resource_path('Hover.mp3')

def play_sound(sound_type):
    """Play the appropriate sound based on type if not muted"""
    if sound_muted:
        return
    try:
        if sound_type == "menu":
            sound_to_play = menu_sound_path if os.path.exists(menu_sound_path) else resource_path('Click.mp3')
        elif sound_type == "main":
            sound_to_play = bing_sound_path if os.path.exists(bing_sound_path) else resource_path('Button.mp3')
        elif sound_type == "Hover":
            sound_to_play = Hover_sound_path if os.path.exists(Hover_sound_path) else resource_path('Hover.mp3')
        
        pygame.mixer.music.load(sound_to_play)
        pygame.mixer.music.play()
    except Exception as e:
        print(f"Error playing sound: {e}")

def toggle_mute():
    """Toggle sound mute state"""
    global sound_muted
    sound_muted = not sound_muted
    # Update the mute button text
    for widget in menu_frame.winfo_children():
        if isinstance(widget, ttk.Button) and widget['text'] in ["Mute", "Unmute"]:
            widget.config(text="Unmute" if sound_muted else "Mute")
            break

def on_enter(event):
    """Handle mouse enter events for main button"""
    if event.widget == button_label:
        event.widget.config(bg='black')

def on_leave(event):
    """Handle mouse leave events for main button"""
    if event.widget == button_label:
        event.widget.config(bg='SystemButtonFace')

def get_monitor(window):
    """Get monitor info where window resides"""
    window_x = window.winfo_rootx()
    window_y = window.winfo_rooty()

    for monitor in get_monitors():
        if monitor.x <= window_x <= monitor.x + monitor.width and monitor.y <= window_y <= monitor.y + monitor.height:
            return monitor
    return get_monitors()[0]

def is_connected():
    """Check internet connectivity"""
    try:
        socket.create_connection(("www.google.com", 80), timeout=5)
        return True
    except OSError:
        return False

def toggle_fullscreen(event=None):
    """Toggle between fullscreen and windowed mode"""
    global fullscreen
    current_monitor = get_monitor(root)

    screen_width = current_monitor.width
    screen_height = current_monitor.height
    monitor_x = current_monitor.x
    monitor_y = current_monitor.y

    if fullscreen:
        root.overrideredirect(False)
        root.geometry("400x300")
        fullscreen = False
    else:
        root.overrideredirect(True)
        root.geometry(f"{screen_width}x{screen_height}+{monitor_x}+{monitor_y}")
        fullscreen = True

def dark_title_bar(window):
    """Apply dark mode title bar (Windows 10/11 only)"""
    window.update()
    DWMWA_USE_IMMERSIVE_DARK_MODE = 20
    set_window_attribute = ct.windll.dwmapi.DwmSetWindowAttribute
    hwnd = ct.windll.user32.GetParent(window.winfo_id())
    value = ct.c_int(2)
    result = set_window_attribute(hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE, ct.byref(value), ct.sizeof(value))
    if result != 0:  # If failed, try with value = 1 (older Windows 10 versions)
        value = ct.c_int(1)
        set_window_attribute(hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE, ct.byref(value), ct.sizeof(value))

def format_text(text):
    """Format text with proper capitalization and punctuation"""
    # Capitalize first letter
    if text:
        text = text[0].upper() + text[1:]
    
    # Add period if no punctuation at end
    if text and text[-1] not in {'.', '!', '?', ',', ';', ':'}:
        text += '.'
    
    return text

def handle_special_commands(text):
    """Handle special voice commands"""
    text = text.lower().strip()
    
    if "scrub" in text:
        if "word" in text:
            pyautogui.hotkey('ctrl', 'backspace')
        elif "last" in text or "that" in text:
            pyautogui.press('backspace')
        elif "all" in text:
            pyautogui.hotkey('ctrl', 'a')
            pyautogui.press('backspace')
        return True
    elif "submit" in text:
        pyautogui.press('enter')
        return True
    elif "tab" in text:
        pyautogui.press('tab')
        return True
    
    return False

def voice_typing():
    """Handle voice typing into any text field"""
    global typing_active, last_typed_position
    
    try:
        with sr.Microphone() as source:
            recognizer.adjust_for_ambient_noise(source)
            while typing_active:
                audio = recognizer.listen(source, phrase_time_limit=5)
                try:
                    text = recognizer.recognize_google(audio)
                    
                    # Handle special commands first
                    if handle_special_commands(text):
                        continue
                        
                    # Format the text
                    formatted_text = format_text(text)
                    
                    # Type the formatted text
                    pyautogui.typewrite(formatted_text + " ")
                    
                except sr.UnknownValueError:
                    pass
                except sr.RequestError as e:
                    print(f"Could not request results; {e}")
    except OSError:
        pass
    finally:
        typing_active = False
        update_listening_indicators()

def start_voice_typing(event=None):
    """Start or stop voice typing"""
    global typing_active, voice_interaction_enabled
    
    if not typing_active:
        play_sound("main")
        typing_active = True
        voice_interaction_enabled = True
        update_listening_indicators()
        threading.Thread(target=voice_typing, daemon=True).start()
    else:
        play_sound("main")
        typing_active = False
        voice_interaction_enabled = False
        update_listening_indicators()

def speak_text(text):
    """Text-to-speech function"""
    try:
        engine.stop()
        engine.say(text)
        engine.runAndWait()
    except Exception as e:
        print(f"Error in speak_text: {e}")

def on_close():
    """Gracefully exit the application"""
    global stop_listening, typing_active
    try:
        typing_active = False
        if stop_listening:
            stop_listening(wait_for_stop=False)
        engine.stop()
        pygame.mixer.quit()
    except Exception as e:
        print(f"Error in on_close: {e}")
    root.destroy()

def center_window_on_parent(child, width, height):
    """Center child window on parent window"""
    parent_x = root.winfo_x()
    parent_y = root.winfo_y()
    parent_width = root.winfo_width()
    parent_height = root.winfo_height()
    
    x = parent_x + (parent_width // 2) - (width // 2)
    y = parent_y + (parent_height // 2) - (height // 2)
    child.geometry(f'{width}x{height}+{x}+{y}')

def show_rules():
    """Display the rules window"""
    play_sound("menu")
    rules_window = tk.Toplevel(root)
    rules_window.title("Rules")
    rules_window.geometry('400x300')
    center_window_on_parent(rules_window, 400, 300)

    rules_label = tk.Label(rules_window, 
                         text="VocalScrub Voice Typing\n\n"
                              "• Click button to start/stop voice typing\n"
                              "• Select text field and speak naturally\n"
                              "• Commands: 'scrub word', 'scrub last',\n"
                              "  'scrub all', 'submit', 'tab'\n\n"
                              "Auto-formats with proper capitalization\n"
                              "and punctuation.\n\n"
                              "A hands-free Windows typing solution.\n\n"
                              "Created by Caleb W. Broussard",
                         font=("Helvetica", 12),
                         wraplength=380,
                         justify='center')  # Changed to center alignment
    rules_label.pack(padx=10, pady=10)

    close_button = tk.Button(rules_window, 
                           text="Close", 
                           command=lambda: [play_sound("menu"), rules_window.destroy()], 
                           font=("Helvetica", 10))
    close_button.pack(pady=5)

def change_wallpaper():
    """Open file explorer to change wallpaper"""
    play_sound("menu")
    file_path = fd.askopenfilename(initialdir=os.getcwd(), title="Select Wallpaper", filetypes=[("Image Files", "*.png;*.jpg;*.jpeg;*.bmp")])
    if file_path:
        try:
            background_image = Image.open(file_path)
            background_photo = ImageTk.PhotoImage(background_image)
            canvas.create_image(0, 0, anchor=tk.NW, image=background_photo)
            canvas.image = background_photo
        except Exception as e:
            print(f"Failed to update wallpaper: {e}")

def change_icon():
    """Open file explorer to change icon"""
    play_sound("menu")
    file_path = fd.askopenfilename(initialdir=os.getcwd(), title="Select Icon", filetypes=[("Image Files", "*.png;*.gif;*.ico")])
    if file_path:
        try:
            if file_path.endswith(".gif"):
                gif_image = Image.open(file_path)
                gif_frames = []
                for frame in range(gif_image.n_frames):
                    gif_image.seek(frame)
                    frame_image = ImageTk.PhotoImage(gif_image.copy())
                    gif_frames.append(frame_image)
                
                def update_gif(frame=0):
                    button_label.config(image=gif_frames[frame])
                    button_label.image = gif_frames[frame]
                    root.after(100, update_gif, (frame + 1) % len(gif_frames))
                
                update_gif()
            else:
                new_button_image = Image.open(file_path)
                new_button_image = new_button_image.resize((150, 150), Image.LANCZOS)
                new_button_photo = ImageTk.PhotoImage(new_button_image)
                button_label.config(image=new_button_photo)
                button_label.image = new_button_photo
        except Exception as e:
            print(f"Failed to update button: {e}")

def change_sound():
    """Change sound settings"""
    play_sound("menu")
    sound_window = tk.Toplevel(root)
    sound_window.title("Sound")
    sound_window.geometry('300x200')
    center_window_on_parent(sound_window, 300, 200)
    
    def select_sound(sound_type):
        play_sound("menu")
        file_path = fd.askopenfilename(initialdir=os.getcwd(), title=f"Select {sound_type} Sound",
                                    filetypes=[("Sound Files", "*.mp3;*.wav")])
        if file_path:
            try:
                global bing_sound_path, menu_sound_path, Hover_sound_path
                if sound_type == "Menu Button":
                    menu_sound_path = file_path
                elif sound_type == "Main Button":
                    bing_sound_path = file_path
                elif sound_type == "Hover":
                    Hover_sound_path = file_path
                play_sound("menu" if sound_type == "Menu Button" else "main")
            except Exception as e:
                print(f"Failed to update sound: {e}")
        sound_window.destroy()
    
    menu_sound_btn = ttk.Button(sound_window, text="Click Sound", command=lambda: select_sound("Menu Button"))
    menu_sound_btn.pack(pady=5)
    main_sound_btn = ttk.Button(sound_window, text="Main Button Sound", command=lambda: select_sound("Main Button"))
    main_sound_btn.pack(pady=5)
    Hover_sound_btn = ttk.Button(sound_window, text="Hover Sound", command=lambda: select_sound("Hover"))
    Hover_sound_btn.pack(pady=5)

def update_listening_indicators():
    """Update all listening indicators based on the listening state"""
    if voice_interaction_enabled or typing_active:
        active_indicator_top_right.lift()
        active_indicator_bottom_right.lift()
        active_indicator_bottom_left.lift()
    else:
        active_indicator_top_right.lower(listening_indicator_top_right)
        active_indicator_bottom_right.lower(listening_indicator_bottom_right)
        active_indicator_bottom_left.lower(listening_indicator_bottom_left)

def change_indicator_colors():
    """Allow user to change indicator colors"""
    play_sound("menu")
    color_window = tk.Toplevel(root)
    color_window.title("Colors")
    color_window.geometry('350x250')
    center_window_on_parent(color_window, 350, 250)
    
    color_examples = [
        "Red", "Green", "Blue", "Yellow", "Purple", 
        "Orange", "Pink", "Cyan", "Magenta", "Lime",
        "DodgerBlue", "Gold", "Violet", "Turquoise", "Salmon",
        "White", "Black", "Gray", "Maroon", "Olive"
    ]
    
    active_frame = tk.Frame(color_window)
    active_frame.pack(pady=5)
    tk.Label(active_frame, text="Active Color:", font=("Helvetica", 10)).pack()
    active_entry = tk.Entry(active_frame, width=20, font=("Helvetica", 10))
    active_entry.insert(0, indicator_colors['active'])
    active_entry.pack(pady=5)
    
    inactive_frame = tk.Frame(color_window)
    inactive_frame.pack(pady=5)
    tk.Label(inactive_frame, text="Inactive Color:", font=("Helvetica", 10)).pack()
    inactive_entry = tk.Entry(inactive_frame, width=20, font=("Helvetica", 10))
    inactive_entry.insert(0, indicator_colors['inactive'])
    inactive_entry.pack(pady=5)
    
    examples_label = tk.Label(color_window, 
                            text=f"Color examples: {', '.join(color_examples)}",
                            font=("Helvetica", 8),
                            wraplength=330)
    examples_label.pack(pady=5)
    
    def apply_colors():
        play_sound("menu")
        new_active = active_entry.get()
        new_inactive = inactive_entry.get()
        
        try:
            test_label = tk.Label(color_window)
            test_label.config(bg=new_active)
            test_label.config(bg=new_inactive)
            
            indicator_colors['active'] = new_active
            indicator_colors['inactive'] = new_inactive
            
            # Update active indicators
            active_indicator_top_right.config(bg=new_active)
            active_indicator_bottom_right.config(bg=new_active)
            active_indicator_bottom_left.config(bg=new_active)
            
            # Update inactive indicators
            listening_indicator_top_right.config(bg=new_inactive)
            listening_indicator_bottom_right.config(bg=new_inactive)
            listening_indicator_bottom_left.config(bg=new_inactive)
            
            color_window.destroy()
        except tk.TclError:
            pass
    
    button_frame = tk.Frame(color_window)
    button_frame.pack(pady=5)
    tk.Button(button_frame, text="Apply", command=apply_colors, font=("Helvetica", 8)).pack(side=tk.LEFT, padx=5)
    tk.Button(button_frame, text="Close", command=lambda: [play_sound("menu"), color_window.destroy()], 
              font=("Helvetica", 8)).pack(side=tk.LEFT, padx=5)

def animate_menu(step, direction):
    """Animate the menu sliding in/out"""
    global menu_animation_id, menu_visible
    
    menu_width = 120
    menu_height = 217  # Reduced height to remove empty space
    
    if direction == "in":
        new_x = -menu_width + (step * 12)
        if new_x >= 0:
            new_x = 0
            menu_visible = True
            menu_button.place_forget()
        else:
            menu_animation_id = root.after(10, animate_menu, step + 1, direction)
    else:
        new_x = 0 - (step * 12)
        if new_x <= -menu_width:
            new_x = -menu_width
            menu_visible = False
            menu_button.place(relx=0, rely=0, x=10, y=10)
        else:
            menu_animation_id = root.after(10, animate_menu, step + 1, direction)
    
    menu_frame.place(x=new_x, y=0, width=menu_width, height=menu_height)
    
    if new_x == 0 or new_x == -menu_width:
        menu_animation_id = None

def toggle_menu():
    """Toggle menu visibility with animation"""
    global menu_animation_id, menu_visible
    
    play_sound("menu")
    
    if menu_animation_id:
        root.after_cancel(menu_animation_id)
        menu_animation_id = None
    
    if menu_visible:
        animate_menu(1, "out")
    else:
        animate_menu(1, "in")

def close_menu_if_open(event):
    """Close menu if clicked outside of it"""
    if menu_visible:
        x, y = event.x, event.y
        if x > menu_frame.winfo_width() or y > menu_frame.winfo_height():
            toggle_menu()

# GUI Initialization
root = tk.Tk()
root.title('VocalScrub')
root.minsize(400, 300)
root.maxsize(1920, 1080)

# Apply dark title bar immediately
root.withdraw()
root.update()
dark_title_bar(root)
root.deiconify()

icon_path = resource_path('Icon.png')
if os.path.exists(icon_path):
    root.iconphoto(True, tk.PhotoImage(file=icon_path))

root.geometry('400x300')
root.update_idletasks()
x = (root.winfo_screenwidth() // 2) - (400 // 2)
y = (root.winfo_screenheight() // 2) - (300 // 2)
root.geometry(f'400x300+{x}+{y}')

# Periodic title bar refresh
def ensure_dark_title():
    dark_title_bar(root)
    root.after(100, ensure_dark_title)
root.after(100, ensure_dark_title)

# Background Setup
background_path = resource_path('Wallpaper.png')
if os.path.exists(background_path):
    try:
        background_image = Image.open(background_path)
        background_photo = ImageTk.PhotoImage(background_image)
        canvas = tk.Canvas(root, width=400, height=300, highlightthickness=0, bd=0)
        canvas.pack(fill="both", expand=True)
        canvas.create_image(0, 0, anchor=tk.NW, image=background_photo)
    except Exception as e:
        print(f"Failed to load background image: {e}")

# Menu Frame
menu_frame = tk.Frame(root, bg='#000000', bd=0)
menu_frame.place(x=-120, y=0, width=120, height=220)  # Reduced height

# Menu Buttons (removed "Type" button and extra spacing)
button_pady = 5
menu_buttons = [
    ("Rules", show_rules),
    ("Colors", change_indicator_colors),
    ("Wallpaper", change_wallpaper),
    ("Icon", change_icon),
    ("Sound", change_sound),
    ("Mute", toggle_mute)
]

for text, command in menu_buttons:
    btn = ttk.Button(
        menu_frame, 
        text=text, 
        command=command, 
        style='Small.TButton'
    )
    btn.pack(pady=button_pady, padx=5, fill=tk.X)
    btn.bind("<Enter>", lambda e: play_sound("Hover"))
    btn.bind("<Leave>", lambda e: None)

style = ttk.Style()
style.configure('Small.TButton', 
               borderwidth=1, 
               padding=(1, 2),
               font=('Helvetica', 8))

# Main Button Setup
image_path = resource_path('Button.png')
if os.path.exists(image_path):
    try:
        button_image = Image.open(image_path)
        button_image = button_image.resize((150, 150), Image.LANCZOS)
        button_image = ImageTk.PhotoImage(button_image)

        button_label = tk.Label(root, image=button_image, bd=0, highlightthickness=0, borderwidth=2, relief=tk.RAISED)
        button_label.place(relx=0.5, rely=0.5, anchor=tk.CENTER, y=-10)
        button_label.bind("<Button-1>", start_voice_typing)
        button_label.bind("<Enter>", on_enter)
        button_label.bind("<Leave>", on_leave)
    except Exception as e:
        print(f"Failed to load button image: {e}")

# Hamburger Menu Button
menu_button = ttk.Button(
    root, 
    text="☰", 
    command=toggle_menu, 
    style='Small.TButton', 
    width=4
)
menu_button.place(relx=0, rely=0, x=10, y=10)

# Listening Indicators - All 2x1 size with proper outlines
# Inactive indicators (black with thin border)
listening_indicator_top_right = tk.Label(root, width=2, height=1, bg='black', bd=1, relief='solid')
listening_indicator_top_right.place(relx=1.0, rely=0, x=-25, y=10)

listening_indicator_bottom_right = tk.Label(root, width=2, height=1, bg='black', bd=1, relief='solid')
listening_indicator_bottom_right.place(relx=1.0, rely=1.0, x=-25, y=-25)

listening_indicator_bottom_left = tk.Label(root, width=2, height=1, bg='black', bd=1, relief='solid')
listening_indicator_bottom_left.place(relx=0, rely=1.0, x=10, y=-25)

# Active indicators (blue with thick black outline)
active_indicator_top_right = tk.Label(root, width=2, height=1, bg='bisque', bd=2, relief='solid')
active_indicator_top_right.place(relx=1.0, rely=0, x=-25, y=10)
active_indicator_top_right.lower(listening_indicator_top_right)

active_indicator_bottom_right = tk.Label(root, width=2, height=1, bg='bisque', bd=2, relief='solid')
active_indicator_bottom_right.place(relx=1.0, rely=1.0, x=-25, y=-25)
active_indicator_bottom_right.lower(listening_indicator_bottom_right)

active_indicator_bottom_left = tk.Label(root, width=2, height=1, bg='bisque', bd=2, relief='solid')
active_indicator_bottom_left.place(relx=0, rely=1.0, x=10, y=-25)
active_indicator_bottom_left.lower(listening_indicator_bottom_left)

# Window Management
root.protocol('WM_DELETE_WINDOW', on_close)
root.bind('<Escape>', toggle_fullscreen)
root.bind('<Button-1>', close_menu_if_open)

root.mainloop()
