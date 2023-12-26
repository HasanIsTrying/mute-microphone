import keyboard
import win32api                                  #app to mute or unmute with a button click
import win32gui
from tkinter import *
from tkinter import messagebox

# -1: muted ___  1: unmuted
is_muted = -1
window = Tk()
recorded_events = []
the_key = 'shift+f2'
hotkey_callback = None
bg_color = '#9b9b9b'
fg_color = '#b80000'
label = Label(window, text='Welcome.\nyou can mute/unmute your microphone here.\n\n\n')
label.config(background=bg_color, foreground=fg_color, font=('Arial', 11, 'bold'))
label.pack()
# Create the button and label once
dis_keyboard_btn = Button(window, text='', font=('Comic Sans', 20),
                          foreground=fg_color, width=17)
dis_keyboard_btn.pack()

label2 = Label(window, text='', background=bg_color)
label2.pack()

#to create a new shortcut listens for all keys pressed until esc is pressed then it saves them
def shortcut():
    global recorded_events, the_key

    messagebox.showinfo(title='How to Create the shortcut!',
                        message='to make a new shortcut to mute/unmute your mic\n'
                                'press (enter) -> press the shortcut you want -> press (esc) to save it')

    # Reset recorded_events before hooking the keyboard
    recorded_events = []
    keyboard.hook(on_key_event)

    try:
        # Wait for 'esc' key without adding it to the recorded events
        keyboard.wait('esc', suppress=True)
    except KeyboardInterrupt:  # Handle the case when 'esc' is pressed
        pass

    keyboard.unhook_all()

    hotkey_sequence = '+'.join(recorded_events)
    print(f"Recorded Hotkey Sequence: {hotkey_sequence}")

    # Save the hotkey sequence without 'esc'
    if_has_esc(hotkey_sequence)
    print(f"Saved Hotkey Sequence: {the_key}")

    try:
        file_hotkey = open('The_shortcut.txt', 'w')
        file_hotkey.write(the_key)
        file_hotkey.close()
    except Exception as e:
        print(f'Error saving hotkey sequence: {e}')

    update_hotkey_listener()

    print(the_key)



def on_key_event(e):
    if e.event_type == keyboard.KEY_DOWN:
        print(f'Key {e.name} pressed')
        recorded_events.append(e.name)

def if_has_esc(key):
    global the_key
    if '+esc' in key:
        the_key = key[:-4]
def update_hotkey_listener():
    global hotkey_callback, the_key
    if hotkey_callback:
        try:
            keyboard.remove_hotkey(hotkey_callback)
        except ValueError:
            # Handle the case where the hotkey is not found
            pass

    try:
        file = open('The_shortcut.txt', 'r')
        the_key = file.readline().strip()
        file.close()
    except:
        print('Error reading hotkey from file')

    if the_key:  # Check if the_key is not empty
        hotkey_callback = keyboard.add_hotkey(the_key, lambda: disable_microphone())



#gui window
def main():
    global window

    window.geometry('350x200')
    window.title('Microphone mute Tool')
    window.config(background=bg_color)

    menubar = Menu(window)
    window.config(menu=menubar)
    settingsbar = Menu(menubar, tearoff=0)
    menubar.add_cascade(menu=settingsbar, label='Settings')
    settingsbar.add_command(label='About', command=read_me_btn)
    settingsbar.add_command(label='Create a new shortcut', command=shortcut)
    settingsbar.add_command(label='show shortcut', command=show_shortcut)
    settingsbar.add_separator()
    settingsbar.add_command(label='Exit', command=window.destroy)

    update_gui()

    window.protocol("WM_DELETE_WINDOW", lambda: window.destroy())
    window.mainloop()

def show_shortcut():  #to show the current shortcut
    messagebox.showinfo(title='your mute shortcut!', message='Here is the shortcut you set: '+the_key)

#to update the window according to if the mic is muted or not
def update_gui():
    global is_muted

    if is_muted == -1:
        button_text = 'Mute Microphone'
        label_text = '\nMicrophone is on\n\n'
    else:
        button_text = 'Unmute Microphone'
        label_text = '\nMicrophone is off\n\n'

    dis_keyboard_btn.config(text=button_text,command=disable_microphone)
    label2.config(text=label_text)

#to disable the microphone
def disable_microphone():
    global is_muted
    WM_APPCOMMAND = 0x319
    APPCOMMAND_MICROPHONE_VOLUME_MUTE = 0x180000

    hwnd_active = win32gui.GetForegroundWindow()
    win32api.SendMessage(hwnd_active, WM_APPCOMMAND, None, APPCOMMAND_MICROPHONE_VOLUME_MUTE)
    is_muted = is_muted * -1

    update_gui()

def read_me_btn():
    messagebox.showinfo(title='Info!', message='This program can be used if you own a microphone that does not have a mute button.\n')

#at the start of the program we read what was the saved shortcut
try:
    file = open('The_shortcut.txt', 'r')
    the_key = file.readline().strip()
    file.close()
except:
    print('empty file')
update_hotkey_listener()
if __name__ == '__main__':
    main()