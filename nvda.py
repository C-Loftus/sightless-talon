from talon import actions, Module, settings, cron, Context, clip, registry
import os, ctypes


mod = Module()
ctx = Context()

mod.tag("nvda_running", desc="NVDA is running")

mod.setting(
    "nvda_key",
    type=str,
    default="capslock",
    desc="The key that is used as the NVDA key",
)

mod.setting(
    "tts_via_screenreader",
    type=bool,
    default=False,
    desc="If true, plays back dictation with text to speech through the screenreader, not within Talon",
)

import subprocess
def is_nvda_running2():
    try:
        # Use the 'tasklist' command to check if NVDA is running
        result = subprocess.run(['tasklist', '/FI', 'IMAGENAME eq nvda.exe'], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)

        # Check if 'nvda.exe' is present in the tasklist output
        return 'nvda.exe' in result.stdout
    except Exception as e:
        print(f"An error occurred: {e}")
        return False



@mod.scope
def check_nvda():
    '''Check if NVDA is running'''
    # Check if it is running without spawning a thread or process
    def is_nvda_running():
        try:
            hwnd = ctypes.windll.user32.FindWindowW(None, "NVDA")
            return hwnd != 0
        except Exception as e:
            print(f"Error: {e}")
            return False

    nvda_running = is_nvda_running()
    if nvda_running:
        ctx.tags = ["user.nvda_running"]
    else:
        ctx.tags = []
    print(list(ctx.tags))

# Re-run the above code every 15s to update the scope
cron.interval("3s", check_nvda.update)



@mod.action_class
class Actions:
     
    def toggle_nvda():
        '''Toggles NVDA on and off'''
        if not actions.user.is_nvda_running():
            actions.key("ctrl-alt-n") 
            actions.user.robot_tts("NVDA on")
        elif actions.user.is_nvda_running():
            actions.user.with_nvda_mod_press('q')
            actions.user.robot_tts("NVDA off")


    def with_nvda_mod_press(key: str):
        """Presses the NVDA key"""
        nvda_key = settings.get("user.nvda_key")
        actions.key(f'{nvda_key}:down') 
        actions.sleep("50ms")
        actions.key(key)
        actions.sleep("10ms")
        actions.key(f'{nvda_key}:up') 

    def is_nvda_running() -> bool:
        '''Returns true if NVDA is running'''
        return "user.nvda_running" in ctx.tags
    
    def nvda_tts(text: str):
        '''text to speech with NVDA'''

ctxNVDARunning = Context()
ctxNVDARunning.matches = r"""
tag: user.nvda_running
"""

@ctxNVDARunning.action_class("user")
class NVDAActions:
    def nvda_tts(text: str):
        """text to speech with NVDA"""
        with clip.revert():
            clip.set_text(text) # sets the result to the clipboard
            actions.sleep("50ms")
            actions.user.with_nvda_mod_press('c')



ctxWindowsNVDARunning = Context()
ctxWindowsNVDARunning.matches = r"""
os: windows
tag: user.nvda_running
"""
@ctxWindowsNVDARunning.action_class('user')
class UserActions:
    
    def robot_tts(text: str):
        """text to speech"""
        
        if settings.get("user.tts_via_screenreader"):
            actions.user.nvda_tts(text)
        else:
            actions.user.windows_robot_tts(text)