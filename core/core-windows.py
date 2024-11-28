import os
from collections import OrderedDict

from talon import Context, actions, settings

if os.name == "nt":
    import pywintypes
    import win32com.client


SVSFDefault = 0
SVSFlagsAsync = 1
SVSFPurgeBeforeSpeak = 2
SVSFIsFilename = 4
SVSFIsXML = 8
SVSFIsNotXML = 16
SVSFPersistXML = 32


class SAPI5:
    """Supports the microsoft speech API version 5."""

    has_volume = True
    has_rate = True
    has_pitch = True
    min_pitch = -10
    max_pitch = 10
    min_rate = -10
    max_rate = 10
    min_volume = 0
    max_volume = 100
    name = "sapi5"
    priority = 101
    system_output = True

    def __init__(self):
        try:
            # self.object = load_com("SAPI.SPVoice")
            self.object = win32com.client.Dispatch("SAPI.SpVoice")

            self._voices = self._available_voices()
        except (pywintypes.com_error, TypeError):
            raise Exception
        self._pitch = 0

    def _available_voices(self):
        _voices = OrderedDict()
        for v in self.object.GetVoices():
            _voices[v.GetDescription()] = v
        return _voices

    def list_voices(self):
        return list(self._voices.keys())

    def get_voice(self):
        return self.object.Voice.GetDescription()

    def set_voice(self, value):
        self.object.Voice = self._voices[value]
        # For some reason SAPI5 does not reset audio after changing the voice
        # By setting the audio device after changing voices seems to fix this
        # This was noted from information at:
        # http://lists.nvaccess.org/pipermail/nvda-dev/2011-November/022464.html

        # Likewise, if this routine is not changed, then the pitch will sound higher than it should be.
        # https://www.autohotkey.com/boards/viewtopic.php?t=33651#p232026
        self.object.AllowAudioOutputFormatChangesOnNextSet = 0
        self.object.AudioOutputStream.Format.Type = 39  # SAFT48kHz16BitStereo
        self.object.AudioOutputStream = self.object.AudioOutputStream
        self.object.AllowAudioOutputFormatChangesOnNextSet = 1

    def get_pitch(self):
        return self._pitch

    def set_pitch(self, value):
        self._pitch = value

    def get_rate(self):
        return self.object.Rate

    def set_rate(self, value):
        self.object.Rate = value

    def get_volume(self):
        return self.object.Volume

    def set_volume(self, value):
        self.object.Volume = value

    def speak(self, text, interrupt=False):
        if interrupt:
            self.silence()
        # We need to do the pitch in XML here
        textOutput = '<pitch absmiddle="%d">%s</pitch>' % (
            round(self._pitch),
            text.replace("<", "&lt;"),
        )
        self.object.Speak(textOutput, SVSFlagsAsync | SVSFIsXML)

    def silence(self):
        self.object.Speak("", SVSFlagsAsync | SVSFPurgeBeforeSpeak)

    def is_active(self):
        if self.object:
            return True
        return False


if os.name == "nt":
    speaker = SAPI5()

ctxWindows = Context()
ctxWindows.matches = r"""
os: windows
"""


@ctxWindows.action_class("user")
class UserActions:
    def base_win_tts(text: str, interrupt: bool):
        """Base function for windows tts. We expose this
        so we can share the speaker object across files. We don't want
        it to get overridden by the other tts functions"""
        speaker.set_rate(settings.get("user.tts_speed"))
        speaker.set_volume(settings.get("user.tts_volume", 50))
        speaker.speak(text, interrupt)

    def tts(text: str, interrupt: bool = True):
        """text to speech with windows voice"""
        # Base interrupt to the base
        actions.user.base_win_tts(text, interrupt)

    def toggle_reader():
        """Toggles the screen reader on and off"""
        actions.user.toggle_nvda()

    def switch_voice():
        """Switches the voice for the screen reader"""
        voices = speaker.list_voices()
        if len(voices) < 2:
            return

        current = speaker.get_voice()
        index = voices.index(current)
        index = (index + 1) % len(voices)
        speaker.set_voice(voices[index])
        actions.user.tts("Switched to " + voices[index])
