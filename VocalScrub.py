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
import traceback

# Global variables
voice_interaction_enabled = False # This is now mostly used to track the active state for indicators
typing_active = False
stop_listening = None # This seems unused in the current voice_typing loop
sound_muted = False
fullscreen = False
menu_visible = False
menu_animation_id = None


# Initialize speech recognition with error handling
recognizer = sr.Recognizer()
try:
    microphone = sr.Microphone()
    print("Microphone initialized successfully")
except OSError as e:
    print(f"Microphone initialization error: {e}")

# Indicator settings
indicator_colors = {
    'active': 'bisque', # Changed to bisque based on user's last script
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

        # Use pygame.mixer.Sound for shorter, overlapping sounds if needed
        # For button clicks, music.load/play is simpler but non-overlapping
        if os.path.exists(sound_to_play):
             pygame.mixer.music.load(sound_to_play)
             pygame.mixer.music.play()
        else:
             print(f"Sound file not found: {sound_to_play}")

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
        event.widget.config(bg='black') # Changed to black based on user's last script

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
    # Fallback to primary monitor if window is not on any known monitor (shouldn't happen)
    return get_monitors()[0]

def is_connected():
    """Check internet connectivity"""
    try:
        # Attempt to create a socket connection to a known host
        socket.create_connection(("www.google.com", 80), timeout=5)
        return True
    except OSError:
        # If connection fails, assume no internet
        return False

# Added a check for Windows OS before calling ctypes
def dark_title_bar(window):
    """Apply dark mode title bar (Windows 10/11 only)"""
    if sys.platform == "win32":
        window.update()
        try:
            DWMWA_USE_IMMERSIVE_DARK_MODE = 20
            set_window_attribute = ct.windll.dwmapi.DwmSetWindowAttribute
            hwnd = ct.windll.user32.GetParent(window.winfo_id())
            # Try with value 2 first (Windows 11)
            value = ct.c_int(2)
            result = set_window_attribute(hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE, ct.byref(value), ct.sizeof(value))
            if result != 0: # If failed, try with value = 1 (older Windows 10 versions)
                 value = ct.c_int(1)
                 set_window_attribute(hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE, ct.byref(value), ct.sizeof(value))
        except Exception as e:
            # Handle potential errors if DWMAPI is not available or call fails
            print(f"Failed to apply dark title bar: {e}")
    else:
        # For non-Windows platforms, do nothing or handle appropriately
        pass


def format_text(text):
    """Format text with proper capitalization and punctuation"""
    if not text:
        return ""

    # Capitalize first letter
    text = text[0].upper() + text[1:]

    # Add a space before punctuation marks if needed (simple heuristic)
    text = text.replace(" .", ".").replace(" ,", ",").replace(" !", "!").replace(" ?", "?").replace(" ;", ";").replace(" :", ":")

    # Ensure space after punctuation before adding period heuristic
    for p in ['.', '!', '?', ',', ';', ':']:
        text = text.replace(p, p + ' ')
    text = ' '.join(text.split()).strip() # Clean up multiple spaces

    # Add period if no ending punctuation (and it's not just a command word like "tab")
    # Added a simple check to not add period to short command-like words
    if len(text.split()) > 1 and text[-1] not in {'.', '!', '?', ',', ';', ':'}:
        text += '.'

    return text

def handle_special_commands(text):
    """Handle special voice commands"""
    # Normalize text for command checking
    normalized_text = text.lower().strip()

    if "scrub" in normalized_text:
        if "word" in normalized_text:
            pyautogui.hotkey('ctrl', 'backspace')
        elif "last" in normalized_text or "that" in normalized_text:
            pyautogui.press('backspace')
        elif "all" in normalized_text:
            pyautogui.hotkey('ctrl', 'a')
            pyautogui.press('backspace')
        return True # Command handled
    elif "submit" in normalized_text:
        pyautogui.press('enter')
        return True # Command handled
    elif "tab" in normalized_text:
        pyautogui.press('tab')
        return True # Command handled

    return False # No command handled

def voice_typing():
    """Handle voice typing into any text field"""
    global typing_active
    
    # Ensure microphone is available before starting
    if 'microphone' not in globals() or microphone is None:
        print("Microphone not initialized. Cannot start voice typing.")
        # Optionally show a message box to the user
        root.after(0, messagebox.showerror, "Error", "Microphone not available. Check your audio devices.")
        typing_active = False
        root.after(0, update_listening_indicators) # Update indicators on main thread
        return

    try:
        with sr.Microphone() as source:
            # Adjust for ambient noise once when starting typing
            recognizer.adjust_for_ambient_noise(source, duration=1) # Reduced duration slightly

            while typing_active:
                print("Voice Typing: Listening...")
                try:
                    # Listen for a phrase with a timeout
                    audio = recognizer.listen(source, timeout=5, phrase_time_limit=10) # Added timeout, increased phrase_time_limit
                except sr.WaitTimeoutError:
                    # If no speech is detected within timeout, continue loop
                    print("Voice Typing: Listen timeout.")
                    continue
                except Exception as e:
                     print(f"Voice Typing: Listen error: {e}")
                     # Could add a small sleep here before continuing
                     time.sleep(0.1)
                     continue


                print("Voice Typing: Recognizing...")
                try:
                    # Recognize speech using Google
                    text = recognizer.recognize_google(audio)
                    print(f"Voice Typing: Heard: \"{text}\"") # Print what was heard

                    # Handle special commands first
                    if handle_special_commands(text):
                        print(f"Voice Typing: Handled command: \"{text}\"")
                        continue # Skip typing if command was handled

                    # Format the text
                    formatted_text = format_text(text)
                    print(f"Voice Typing: Typing: \"{formatted_text}\"") # Print what will be typed

                    # Type the formatted text
                    pyautogui.typewrite(formatted_text + " ")

                except sr.UnknownValueError:
                    # Speech was unintelligible
                    print("Voice Typing: Could not understand audio")
                    # Optionally provide feedback to the user via TTS
                    # speak_text("Sorry, could not understand that.")
                except sr.RequestError as e:
                    # API was unreachable or unresponsive
                    print(f"Voice Typing: Could not request results from Google Speech Recognition service; {e}")
                    # Optionally provide feedback to the user via TTS
                    # speak_text("Speech service error.")
                    # This is a more serious error, might want to break the loop or wait longer
                    typing_active = False # Stop typing on critical error
                    root.after(0, messagebox.showerror, "Error", f"Speech service error: {e}")
                except Exception as e:
                    # Catch any other unexpected errors during recognition or typing
                    print(f"Voice Typing: An unexpected error occurred: {e}")
                    traceback.print_exc() # Print traceback for debugging
                    typing_active = False # Stop typing on unexpected error


    except OSError as e:
        print(f"Voice Typing: Microphone OSError: {e}")
        root.after(0, messagebox.showerror, "Error", f"Microphone error: {e}")
    except Exception as e:
        print(f"Voice Typing: An unexpected error occurred in the main loop: {e}")
        traceback.print_exc()
    finally:
        print("Voice Typing: Stopping.")
        typing_active = False
        # Use root.after to safely update GUI from the thread
        root.after(0, update_listening_indicators)


# Added a flag check to prevent multiple threads
typing_thread = None

def start_voice_typing(event=None):
    """Start or stop voice typing"""
    global typing_active, voice_interaction_enabled, typing_thread

    # Prevent starting multiple threads
    if typing_thread is not None and typing_thread.is_alive():
         print("Voice typing thread already running.")
         # Assuming the button toggles, just stop the current one
         typing_active = False
         play_sound("main")
         # Indicators will be updated when the thread finishes
         return

    if not typing_active:
        print("Starting voice typing...")
        play_sound("main")
        typing_active = True
        voice_interaction_enabled = True # Keep this linked for indicator logic
        update_listening_indicators()
        # Start the thread
        typing_thread = threading.Thread(target=voice_typing, daemon=True)
        typing_thread.start()
    else:
        print("Stopping voice typing...")
        play_sound("main")
        typing_active = False # Signal the thread to stop
        voice_interaction_enabled = False # Keep this linked for indicator logic
        # Indicators will be updated by the thread's finally block


def speak_text(text):
    """Text-to-speech function"""
    try:
        # Ensure engine is not busy before stopping/saying
        if engine._inLoop:
             engine.endLoop() # Attempt to end previous speak loop if active
        engine.stop() # Ensure it's stopped
        engine.say(text)
        engine.runAndWait() # Blocking call
    except Exception as e:
        print(f"Error in speak_text: {e}")

def on_close():
    """Gracefully exit the application"""
    global typing_active, stop_listening, typing_thread

    print("Closing application...")
    try:
        # Signal voice typing thread to stop
        typing_active = False
        if typing_thread and typing_thread.is_alive():
            print("Joining typing thread...")
            typing_thread.join(timeout=1) # Give thread a moment to finish

        # Stop speech recognition listening if it was manually started (less relevant for voice_typing)
        if stop_listening:
            print("Stopping speech recognition listener...")
            stop_listening(wait_for_stop=False)

        # Stop TTS engine
        if engine:
            print("Stopping TTS engine...")
            engine.stop()
            # Attempt to end engine loop if it's running (might prevent hang)
            if engine._inLoop:
                 try:
                     engine.endLoop()
                 except Exception as e:
                      print(f"Error ending TTS loop: {e}")


        # Stop Pygame mixer
        if pygame.mixer.get_init():
            print("Quitting Pygame mixer...")
            pygame.mixer.quit()

        print("Destroying root window...")
        root.destroy()
        print("Application closed.")

    except Exception as e:
        print(f"Error during application close: {e}")
        traceback.print_exc() # Print traceback for closing errors
        # Force destroy if cleanup fails
        try:
             root.destroy()
        except:
             pass # Ignore if root is already destroyed


def center_window_on_parent(child, width, height):
    """Center child window on parent window"""
    # Make sure parent window is mapped and dimensions are valid
    root.update_idletasks()
    parent_x = root.winfo_x()
    parent_y = root.winfo_y()
    parent_width = root.winfo_width()
    parent_height = root.winfo_height()

    # Avoid division by zero or negative sizes
    if parent_width <= 0 or parent_height <= 0 or width <= 0 or height <= 0:
         print("Warning: Invalid dimensions for centering.")
         return

    x = parent_x + (parent_width // 2) - (width // 2)
    y = parent_y + (parent_height // 2) - (height // 2)
    child.geometry(f'{width}x{height}+{x}+{y}')

def show_rules():
    """Display the rules window"""
    play_sound("menu")
    rules_window = tk.Toplevel(root)
    rules_window.title("Rules")
    rules_window.geometry('400x300')
    # Use root.after to call centering after window is fully created/mapped
    root.after(10, center_window_on_parent, rules_window, 400, 300)

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
                         justify='center')
    rules_label.pack(padx=10, pady=10)

    close_button = tk.Button(rules_window,
                           text="Close",
                           command=lambda: [play_sound("menu"), rules_window.destroy()],
                           font=("Helvetica", 10))
    close_button.pack(pady=5)

def change_wallpaper():
    """Open file explorer to change wallpaper"""
    play_sound("menu")
    file_path = fd.askopenfilename(initialdir=os.getcwd(), title="Select Wallpaper", filetypes=[("Image Files", ".png;.jpg;.jpeg;.bmp")])
    if file_path:
        try:
            # Load the image and resize/position as needed
            background_image = Image.open(file_path)
            # You might want to resize this to fit the canvas or window size
            # Example: background_image = background_image.resize((canvas.winfo_width(), canvas.winfo_height()), Image.LANCZOS)
            background_photo = ImageTk.PhotoImage(background_image)

            # Clear previous background
            canvas.delete("background") # Assuming you use a tag for the background image

            # Draw new background image on canvas
            canvas.create_image(0, 0, anchor=tk.NW, image=background_photo, tags="background")
            canvas.image = background_photo # Keep a reference

            # Ensure other widgets are on top (if using canvas.create_window)
            # Or ensure the image is below other items if they are packed/placed on the root/canvas
            canvas.tag_lower("background") # Move the background image tag to the bottom

        except Exception as e:
            print(f"Failed to update wallpaper: {e}")
            messagebox.showerror("Error", f"Failed to update wallpaper: {e}")


def change_icon():
    """Open file explorer to change icon"""
    play_sound("menu")
    file_path = fd.askopenfilename(initialdir=os.getcwd(), title="Select Icon", filetypes=[("Image Files", ".png;*.gif;.ico")])
    if file_path:
        try:
            if file_path.lower().endswith(".gif"):
                # Handle GIFs (basic animation loop)
                gif_image = Image.open(file_path)
                gif_frames = []
                try:
                    for frame in range(gif_image.n_frames):
                        gif_image.seek(frame)
                        # Resize each frame
                        frame_image = gif_image.copy().resize((150, 150), Image.LANCZOS)
                        gif_frames.append(ImageTk.PhotoImage(frame_image))
                except EOFError:
                     pass # Ignore EOFError in PIL seeking frames

                if gif_frames:
                    def update_gif(frame=0):
                        if button_label.winfo_exists(): # Check if label still exists
                            button_label.config(image=gif_frames[frame])
                            button_label.image = gif_frames[frame] # Keep reference
                            # Schedule next frame
                            root.after(gif_image.info.get('duration', 100), update_gif, (frame + 1) % len(gif_frames))
                    # Start animation
                    update_gif()
                else:
                    print("Could not load GIF frames.")
                    messagebox.showwarning("Warning", "Could not load GIF animation.")

            else:
                # Handle static images
                new_button_image = Image.open(file_path)
                new_button_image = new_button_image.resize((150, 150), Image.LANCZOS)
                new_button_photo = ImageTk.PhotoImage(new_button_image)
                button_label.config(image=new_button_photo)
                button_label.image = new_button_photo # Keep reference

            # Update window icon if it's an .ico file (Windows specific)
            if file_path.lower().endswith(".ico"):
                 root.iconphoto(True, tk.PhotoImage(file=file_path))
                 # For Windows taskbar icon, may need win32api/win32gui or PyInstaller options

        except Exception as e:
            print(f"Failed to update button/icon: {e}")
            messagebox.showerror("Error", f"Failed to update button/icon: {e}")


def change_sound():
    """Change sound settings"""
    play_sound("menu")
    sound_window = tk.Toplevel(root)
    sound_window.title("Sound")
    sound_window.geometry('300x200')
    root.after(10, center_window_on_parent, sound_window, 300, 200) # Center after window creation

    def select_sound(sound_type):
        play_sound("menu") # Play sound on button click in sound window
        file_path = fd.askopenfilename(initialdir=os.getcwd(), title=f"Select {sound_type} Sound",
                                    filetypes=[("Sound Files", ".mp3;.wav")])
        if file_path:
            try:
                global bing_sound_path, menu_sound_path, Hover_sound_path
                if sound_type == "Menu Button":
                    menu_sound_path = file_path
                elif sound_type == "Main Button":
                    bing_sound_path = file_path
                elif sound_type == "Hover":
                    Hover_sound_path = file_path

                # Test the new sound immediately
                if sound_type == "Menu Button":
                    play_sound("menu")
                elif sound_type == "Main Button":
                    play_sound("main")
                elif sound_type == "Hover":
                    # Playing a hover sound might be annoying, maybe skip testing this one
                     pass

            except Exception as e:
                print(f"Failed to update sound path: {e}")
                messagebox.showerror("Error", f"Failed to update sound: {e}")
        # Decide if you want to close the window automatically after selecting a sound
        # sound_window.destroy() # Uncomment to close automatically

    menu_sound_btn = ttk.Button(sound_window, text="Menu Click Sound", command=lambda: select_sound("Menu Button"))
    menu_sound_btn.pack(pady=5)
    main_sound_btn = ttk.Button(sound_window, text="Main Button Sound", command=lambda: select_sound("Main Button"))
    main_sound_btn.pack(pady=5)
    Hover_sound_btn = ttk.Button(sound_window, text="Hover Sound", command=lambda: select_sound("Hover"))
    Hover_sound_btn.pack(pady=5)

    # Add a close button if not closing automatically
    close_btn = ttk.Button(sound_window, text="Close", command=lambda: [play_sound("menu"), sound_window.destroy()])
    close_btn.pack(pady=5)


def update_listening_indicators():
    """Update all listening indicators based on the listening state"""
    # Indicator state is now solely based on typing_active flag
    if typing_active:
        # Use configured active color
        active_indicator_top_right.config(bg=indicator_colors['active'], bd=2)
        active_indicator_bottom_right.config(bg=indicator_colors['active'], bd=2)
        active_indicator_bottom_left.config(bg=indicator_colors['active'], bd=2)

        active_indicator_top_right.lift()
        active_indicator_bottom_right.lift()
        active_indicator_bottom_left.lift()
    else:
         # Use configured inactive color
        active_indicator_top_right.config(bg=indicator_colors['inactive'], bd=1)
        active_indicator_bottom_right.config(bg=indicator_colors['inactive'], bd=1)
        active_indicator_bottom_left.config(bg=indicator_colors['inactive'], bd=1)

        active_indicator_top_right.lower(listening_indicator_top_right)
        active_indicator_bottom_right.lower(listening_indicator_bottom_right)
        active_indicator_bottom_left.lower(listening_indicator_bottom_left)

    # Also update the inactive indicators to match the inactive color setting
    listening_indicator_top_right.config(bg=indicator_colors['inactive'], bd=1)
    listening_indicator_bottom_right.config(bg=indicator_colors['inactive'], bd=1)
    listening_indicator_bottom_left.config(bg=indicator_colors['inactive'], bd=1)


def change_indicator_colors():
    """Allow user to change indicator colors"""
    play_sound("menu")
    color_window = tk.Toplevel(root)
    color_window.title("Colors")
    color_window.geometry('350x250')
    root.after(10, center_window_on_parent, color_window, 350, 250) # Center after creation

    color_examples = [
        "Red", "Green", "Blue", "Yellow", "Purple",
        "Orange", "Pink", "Cyan", "Magenta", "Lime",
        "DodgerBlue", "Gold", "Violet", "Turquoise", "Salmon",
        "White", "Black", "Gray", "Maroon", "Olive", "Bisque", "Azure", "Navy" # Added some you used
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
        new_active = active_entry.get().strip()
        new_inactive = inactive_entry.get().strip()

        # Simple validation: try setting a temporary label's background
        try:
            test_label = tk.Label(color_window)
            test_label.config(bg=new_active)
            test_label.config(bg=new_inactive)
            test_label.destroy() # Clean up temp label
        except tk.TclError:
            messagebox.showerror("Invalid Color", "One or both colors are invalid.")
            return # Stop if colors are invalid

        # If valid, update global settings and indicators
        indicator_colors['active'] = new_active
        indicator_colors['inactive'] = new_inactive

        # Update indicators immediately
        update_listening_indicators()

        color_window.destroy()

    button_frame = tk.Frame(color_window)
    button_frame.pack(pady=5)
    tk.Button(button_frame, text="Apply", command=apply_colors, font=("Helvetica", 8)).pack(side=tk.LEFT, padx=5)
    tk.Button(button_frame, text="Close", command=lambda: [play_sound("menu"), color_window.destroy()],
              font=("Helvetica", 8)).pack(side=tk.LEFT, padx=5)

def animate_menu(step, direction):
    """Animate the menu sliding in/out"""
    global menu_animation_id, menu_visible

    # Adjust menu_width if needed, keeping it constant during animation
    menu_width = menu_frame.winfo_width() if menu_frame.winfo_width() > 1 else 120 # Use current width or default

    if direction == "in":
        # Calculate new x position, sliding from left (-menu_width) to 0
        new_x = -menu_width + (step * 12) # Increased step for slightly faster animation
        if new_x >= 0:
            new_x = 0
            menu_visible = True
            if menu_button.winfo_exists(): menu_button.place_forget() # Hide hamburger button
        else:
            menu_animation_id = root.after(10, animate_menu, step + 1, direction) # Schedule next step
    else: # direction == "out"
        # Calculate new x position, sliding from 0 to -menu_width
        new_x = 0 - (step * 12) # Increased step
        if new_x <= -menu_width:
            new_x = -menu_width
            menu_visible = False
            if menu_button.winfo_exists(): menu_button.place(relx=0, rely=0, x=10, y=10) # Show hamburger button
        else:
            menu_animation_id = root.after(10, animate_menu, step + 1, direction) # Schedule next step

    # Place the menu frame at the calculated x position
    menu_frame.place(x=new_x, y=0, width=menu_width, height=menu_frame.winfo_height() if menu_frame.winfo_height() > 1 else 220) # Use current height or default

    # If animation is complete, clear the animation ID
    if new_x == 0 or new_x == -menu_width:
        menu_animation_id = None


def toggle_menu():
    """Toggle menu visibility with animation"""
    global menu_animation_id, menu_visible

    play_sound("menu")

    # Cancel any ongoing animation first
    if menu_animation_id:
        root.after_cancel(menu_animation_id)
        menu_animation_id = None

    # Determine direction and start animation
    if menu_visible:
        animate_menu(1, "out")
    else:
        animate_menu(1, "in")


def close_menu_if_open(event):
    """Close menu if clicked outside of it"""
    # Check if menu is visible AND a menu animation is NOT running
    if menu_visible and menu_animation_id is None:
        # Get event coordinates relative to the root window
        x, y = event.x, event.y

        # Get menu frame's current geometry relative to root
        menu_x = menu_frame.winfo_x()
        menu_y = menu_frame.winfo_y()
        menu_width = menu_frame.winfo_width()
        menu_height = menu_frame.winfo_height()

        # Check if click is outside the menu frame boundaries
        if x < menu_x or x > menu_x + menu_width or y < menu_y or y > menu_y + menu_height:
            toggle_menu() # Close the menu


# GUI Initialization
root = tk.Tk()
root.title('VocalScrub')
root.minsize(400, 300)
root.maxsize(1920, 1080) # Keep maxsize reasonable or remove it if truly full screen possible

# Apply dark title bar immediately (Windows only)
root.withdraw() # Hide window during setup
root.update()   # Force update to get window info
dark_title_bar(root) # Apply dark title bar
root.deiconify() # Show window again

icon_path = resource_path('Icon.png')
# Load default icon first
default_icon = None
if os.path.exists(icon_path):
    try:
        default_icon = tk.PhotoImage(file=icon_path)
        root.iconphoto(True, default_icon)
    except Exception as e:
        print(f"Failed to load default icon: {e}")


# Set initial window position and size
initial_width = 400
initial_height = 300
root.geometry(f'{initial_width}x{initial_height}')
root.update_idletasks() # Update geometry information

# Calculate center position
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
x = (screen_width // 2) - (initial_width // 2)
y = (screen_height // 2) - (initial_height // 2)
root.geometry(f'{initial_width}x{initial_height}+{x}+{y}')


# Periodic title bar refresh - REMOVED LOOP as discussed
# The initial call to dark_title_bar should be sufficient on launch
# If title bar reverts later (e.g. on focus change), you might need a bind like <FocusIn>
# root.after(100, ensure_dark_title) # <-- Removed this line

# Background Setup
# Using a Canvas to place background and then other widgets (like button_label) might be cleaner
# than packing things directly on root if you want the background to truly fill behind them.
# If you keep packing/placing directly on root, background image might be covered.
# Let's assume canvas setup from previous versions is intended.

background_path = resource_path('Wallpaper.png')
canvas = tk.Canvas(root, highlightthickness=0, bd=0)
canvas.pack(fill="both", expand=True)
canvas.background_photo = None # Keep a reference to avoid garbage collection

if os.path.exists(background_path):
    try:
        background_image = Image.open(background_path)
        # Resize background image to fit the current window size (or canvas size)
        # This needs to be updated if the window is resized
        # For simplicity on initial load, use initial window size or let canvas handle it
        # A bind on <Configure> could handle resizing later
        bg_width = root.winfo_width()
        bg_height = root.winfo_height()
        if bg_width > 1 and bg_height > 1: # Ensure sizes are valid
             background_image = background_image.resize((bg_width, bg_height), Image.LANCZOS)

        background_photo = ImageTk.PhotoImage(background_image)
        canvas.create_image(0, 0, anchor=tk.NW, image=background_photo, tags="background")
        canvas.background_photo = background_photo # Keep reference
        canvas.tag_lower("background") # Ensure background is at the bottom

    except Exception as e:
        print(f"Failed to load background image: {e}")
        messagebox.showerror("Error", f"Failed to load background image: {e}")

# Menu Frame - Place on root directly, animation changes its x position
menu_frame = tk.Frame(root, bg='#000000', bd=0)
# Place it off-screen initially. Width and height should match actual button container.
# Let's explicitly set size here for clarity.
menu_frame_width = 120
menu_frame_height = 220 # Adjusted height based on user's last script
menu_frame.place(x=-menu_frame_width, y=0, width=menu_frame_width, height=menu_frame_height)

# Menu Buttons
button_pady = 5 # Adjusted pady based on user's last script
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
    # Using lambda within loop requires capturing the current value of 'text'
    #btn.bind("<Enter>", lambda e: play_sound("Hover")) # Original line, works for constant sound
    # If you wanted different hover sounds per button, you'd need more logic

# Apply style (should be done after creating buttons that use it)
style = ttk.Style()
style.configure('Small.TButton',
               borderwidth=1,
               padding=(1, 1),
               font=('Helvetica', 9))


# Main Button Setup
image_path = resource_path('Button.png')
# Load default button image first
button_photo = None
if os.path.exists(image_path):
    try:
        button_image = Image.open(image_path)
        button_image = button_image.resize((150, 150), Image.LANCZOS)
        button_photo = ImageTk.PhotoImage(button_image)
    except Exception as e:
        print(f"Failed to load default button image: {e}")


# Create button_label on the canvas if using canvas for background
if 'canvas' in globals() and canvas is not None:
    button_label = tk.Label(canvas, image=button_photo, bd=0, highlightthickness=0, borderwidth=2, relief=tk.RAISED)
    # Place button relative to canvas center
    button_label.place(relx=0.5, rely=0.5, anchor=tk.CENTER, y=-10) # Adjust y to account for potential status bar later

else:
    # Fallback: Create button_label directly on root if canvas isn't used
    button_label = tk.Label(root, image=button_photo, bd=0, highlightthickness=0, borderwidth=2, relief=tk.RAISED)
    # Place button relative to root center
    button_label.place(relx=0.5, rely=0.5, anchor=tk.CENTER, y=-10)


# Store reference to button_photo to prevent garbage collection
button_label.image = button_photo

# Bind events to the button_label
button_label.bind("<Button-1>", start_voice_typing)
button_label.bind("<Enter>", on_enter)
button_label.bind("<Leave>", on_leave)


# Hamburger Menu Button
menu_button = ttk.Button(
    root,
    text="☰",
    command=toggle_menu,
    style='Small.TButton',
    width=2
)
menu_button.place(relx=0, rely=0, x=10, y=10)


# Listening Indicators - All 2x1 size with proper outlines
# Inactive indicators (colored with thin border) - These are the base labels
listening_indicator_top_right = tk.Label(root, width=2, height=1, bg=indicator_colors['inactive'], bd=1, relief='solid')
listening_indicator_top_right.place(relx=1.0, rely=0, x=-25, y=10)

listening_indicator_bottom_right = tk.Label(root, width=2, height=1, bg=indicator_colors['inactive'], bd=1, relief='solid')
listening_indicator_bottom_right.place(relx=1.0, rely=1.0, x=-25, y=-25)

listening_indicator_bottom_left = tk.Label(root, width=2, height=1, bg=indicator_colors['inactive'], bd=1, relief='solid')
listening_indicator_bottom_left.place(relx=0, rely=1.0, x=10, y=-25)

# Active indicators (colored with thicker border) - These labels are raised when active
active_indicator_top_right = tk.Label(root, width=2, height=1, bg=indicator_colors['active'], bd=2, relief='solid')
active_indicator_top_right.place(relx=1.0, rely=0, x=-25, y=10)
active_indicator_top_right.lower(listening_indicator_top_right) # Start lowered

active_indicator_bottom_right = tk.Label(root, width=2, height=1, bg=indicator_colors['active'], bd=2, relief='solid')
active_indicator_bottom_right.place(relx=1.0, rely=1.0, x=-25, y=-25)
active_indicator_bottom_right.lower(listening_indicator_bottom_right) # Start lowered

active_indicator_bottom_left = tk.Label(root, width=2, height=1, bg=indicator_colors['active'], bd=2, relief='solid')
active_indicator_bottom_left.place(relx=0, rely=1.0, x=10, y=-25)
active_indicator_bottom_left.lower(listening_indicator_bottom_left) # Start lowered

# Initial update of indicators based on starting state (should be inactive)
update_listening_indicators()


# Window Management and Event Bindings

def toggle_fullscreen(event=None):
    """Toggle fullscreen mode for the root window."""
    global fullscreen
    fullscreen = not fullscreen
    root.attributes("-fullscreen", fullscreen)
    # Optionally, you can also update the Escape binding to exit fullscreen
    if fullscreen:
        root.bind('<Escape>', toggle_fullscreen)
    else:
        root.bind('<Escape>', toggle_fullscreen)

root.protocol('WM_DELETE_WINDOW', on_close) # Bind close button to on_close function
root.bind('<Escape>', toggle_fullscreen) # Bind Escape key to fullscreen toggle
root.bind('<Button-1>', close_menu_if_open) # Bind mouse click to close menu


# Start the Tkinter main loop
root.mainloop()
