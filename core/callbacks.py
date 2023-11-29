from talon import scope, registry, ui, actions, settings, speech_system, app

def on_phrase(parsed_phrase):
    if actions.speech.enabled() and settings.get('user.echo_dictation'):
        words = parsed_phrase.get('text')
        if words:
            command_chain = ' '.join(words)

            # Not all tts engines support canceling
            # Easier to just catch the exception
            try:
                actions.user.cancel_robot_tts()
            except:
                pass

            actions.user.robot_tts(command_chain)
speech_system.register('phrase', on_phrase)

def on_app_switch(app):
    if not settings.get("user.echo_context"):
        return 
    actions.user.echo_context()

# We have to keep track of the last title so we don't repeat it
# since sometimes Talon triggers a "title switch" when 
# the title actually hasn't changed, i.e. when a text file is saved
last_title = None
def on_title_switch(win):
    if not settings.get("user.echo_context"):
        return
    window = ui.active_window()
    active_window_title = window.title
    # get just the first two word
    active_window_title = ' '.join(active_window_title.split()[:2])
    #trime the title to 20 characters so super long addresses don't get read
    active_window_title = active_window_title[:20]

    global last_title
    if last_title == active_window_title:
        return
    else:
        last_title = active_window_title

    actions.user.robot_tts(f"{active_window_title}")

last_mode = None
def on_update_contexts():
    global last_mode
    modes = scope.get("mode") or []
    if last_mode == 'sleep' and 'sleep' not in modes:
        actions.user.robot_tts(f'Talon has waken up')
    last_mode = "sleep" if "sleep" in modes else "other"
        
def on_ready():
    if not settings.get("user.start_screenreader_on_startup"):
        return
    actions.user.toggle_nvda()
    actions.user.robot_tts("Talon user scripts loaded")

app.register("ready", on_ready)
registry.register("update_contexts", on_update_contexts)
ui.register("app_activate", on_app_switch)
ui.register("win_title", on_title_switch)