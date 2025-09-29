# noScribe - AI-powered Audio Transcription
# Copyright (C) 2025 Kai Dröge
# ported to MAC by Philipp Schneider (gernophil)

# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.

# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.

# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <https://www.gnu.org/licenses/>.

import sys
# In the compiled version (no command line), stdout is None which might lead to errors
if sys.stdout is None:
    sys.stdout = open(os.devnull, "w")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w")

import tkinter as tk
import customtkinter as ctk
from tkHyperlinkManager import HyperlinkManager
import webbrowser
from functools import partial
from PIL import Image
import os
import platform
import yaml
import locale
import appdirs
from subprocess import run, call, Popen, PIPE, STDOUT
if platform.system() == 'Windows':
    # import torch.cuda # to check with torch.cuda.is_available()
    from subprocess import STARTUPINFO, STARTF_USESHOWWINDOW
if platform.system() in ("Windows", "Linux"):
    from ctranslate2 import get_cuda_device_count
    import torch
import re
if platform.system() == "Darwin": # = MAC
    from subprocess import check_output
    if platform.machine() == "x86_64":
        os.environ['KMP_DUPLICATE_LIB_OK']='True' # prevent OMP: Error #15: Initializing libomp.dylib, but found libiomp5.dylib already initialized.
    # import torch.backends.mps # loading torch modules leads to segmentation fault later
from faster_whisper.audio import decode_audio
from faster_whisper.vad import VadOptions, get_speech_timestamps
import AdvancedHTMLParser
import html
from threading import Thread
import time
from tempfile import TemporaryDirectory
import datetime
from pathlib import Path
if platform.system() in ("Darwin", "Linux"):
    import shlex
if platform.system() == 'Windows':
    import cpufeature
if platform.system() == 'Darwin':
    import Foundation
import logging
import json
import urllib
import multiprocessing
import gc
import traceback

 # Pyinstaller fix, used to open multiple instances on Mac
multiprocessing.freeze_support()

logging.basicConfig()
logging.getLogger("faster_whisper").setLevel(logging.DEBUG)

app_version = '0.6.2'
app_year = '2025'
app_dir = os.path.abspath(os.path.dirname(__file__))

ctk.set_appearance_mode('dark')
ctk.set_default_color_theme('blue')

default_html = """
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0//EN" "http://www.w3.org/TR/REC-html40/strict.dtd">
<html >
<head >
<meta charset="UTF-8" />
<meta name="qrichtext" content="1" />
<style type="text/css" >
p, li { white-space: pre-wrap; }
</style>
<style type="text/css" > 
 p { font-size: 0.9em; } 
 .MsoNormal { font-family: "Arial"; font-weight: 400; font-style: normal; font-size: 0.9em; }
 @page WordSection1 {mso-line-numbers-restart: continuous; mso-line-numbers-count-by: 1; mso-line-numbers-start: 1; }
 div.WordSection1 {page:WordSection1;} 
</style>
</head>
<body style="font-family: 'Arial'; font-weight: 400; font-style: normal" >
</body>
</html>"""

languages = {
    "Auto": "auto",
    "Multilingual": "multilingual",
    "Afrikaans": "af",
    "Arabic": "ar",
    "Armenian": "hy",
    "Azerbaijani": "az",
    "Belarusian": "be",
    "Bosnian": "bs",
    "Bulgarian": "bg",
    "Catalan": "ca",
    "Chinese": "zh",
    "Croatian": "hr",
    "Czech": "cs",
    "Danish": "da",
    "Dutch": "nl",
    "English": "en",
    "Estonian": "et",
    "Finnish": "fi",
    "French": "fr",
    "Galician": "gl",
    "German": "de",
    "Greek": "el",
    "Hebrew": "he",
    "Hindi": "hi",
    "Hungarian": "hu",
    "Icelandic": "is",
    "Indonesian": "id",
    "Italian": "it",
    "Japanese": "ja",
    "Kannada": "kn",
    "Kazakh": "kk",
    "Korean": "ko",
    "Latvian": "lv",
    "Lithuanian": "lt",
    "Macedonian": "mk",
    "Malay": "ms",
    "Marathi": "mr",
    "Maori": "mi",
    "Nepali": "ne",
    "Norwegian": "no",
    "Persian": "fa",
    "Polish": "pl",
    "Portuguese": "pt",
    "Romanian": "ro",
    "Russian": "ru",
    "Serbian": "sr",
    "Slovak": "sk",
    "Slovenian": "sl",
    "Spanish": "es",
    "Swahili": "sw",
    "Swedish": "sv",
    "Tagalog": "tl",
    "Tamil": "ta",
    "Thai": "th",
    "Turkish": "tr",
    "Ukrainian": "uk",
    "Urdu": "ur",
    "Vietnamese": "vi",
    "Welsh": "cy",
}

# config
config_dir = appdirs.user_config_dir('noScribe')
if not os.path.exists(config_dir):
    os.makedirs(config_dir)

config_file = os.path.join(config_dir, 'config.yml')

try:
    with open(config_file, 'r') as file:
        config = yaml.safe_load(file)
        if not config:
            raise # config file is empty (None)        
except: # seems we run it for the first time and there is no config file
    config = {}
    
def get_config(key: str, default):
    """ Get a config value, set it if it doesn't exist """
    if key not in config:
        config[key] = default
    return config[key]

    
def version_higher(version1, version2) -> int:
    """Will return 
    1 if version1 is higher
    2 if version2 is higher
    0  if both are equal """
    version1_elems = version1.split('.')
    version2_elems = version2.split('.')
    # make both versions the same length
    elem_num = max(len(version1_elems), len(version2_elems))
    while len(version1_elems) < elem_num:
        version1_elems.append('0')
    while len(version1_elems) < elem_num:
        version1_elems.append('0')
    for i in range(elem_num):
        if int(version1_elems[i]) > int(version2_elems[i]):
            return 1
        elif int(version2_elems[i]) > int(version1_elems[i]):
            return 2
    # must be completly equal
    return 0
    
# In versions < 0.4.5 (Windows/Linux only), 'pyannote_xpu' was always set to 'cpu'.
# Delete this so we can determine the optimal value  
if platform.system() in ('Windows', 'Linux'):
    try:
        if version_higher('0.4.5', config['app_version']) == 1:
            del config['pyannote_xpu'] 
    except:
        pass

config['app_version'] = app_version

def save_config():
    with open(config_file, 'w') as file:
        yaml.safe_dump(config, file)

save_config()

# locale: setting the language of the UI
# see https://pypi.org/project/python-i18n/
import i18n
from i18n import t
i18n.set('filename_format', '{locale}.{format}')
i18n.load_path.append(os.path.join(app_dir, 'trans'))

try:
    app_locale = config['locale']
except:
    app_locale = 'auto'

if app_locale == 'auto': # read system locale settings
    try:
        if platform.system() == 'Windows':
            app_locale = locale.getdefaultlocale()[0][0:2]
        elif platform.system() == "Darwin": # = MAC
            app_locale = Foundation.NSUserDefaults.standardUserDefaults().stringForKey_('AppleLocale')[0:2]
    except:
        app_locale = 'en'
i18n.set('fallback', 'en')
i18n.set('locale', app_locale)
config['locale'] = app_locale

# determine optimal number of threads for faster-whisper (depending on cpu cores)
if platform.system() == 'Windows':
    number_threads = get_config('threads', cpufeature.CPUFeature["num_physical_cores"])
elif platform.system() == "Linux":
    number_threads = get_config('threads', os.cpu_count() if os.cpu_count() is not None else 4)
elif platform.system() == "Darwin": # = MAC
    if platform.machine() == "arm64":
        cpu_count = int(check_output(["sysctl", "-n", "hw.perflevel0.logicalcpu_max"]))
    elif platform.machine() == "x86_64":
        cpu_count = int(check_output(["sysctl", "-n", "hw.logicalcpu_max"]))
    else:
        raise Exception("Unsupported mac")
    number_threads = get_config('threads', int(cpu_count * 0.75))
else:
    raise Exception('Platform not supported yet.')

# timestamp regex
timestamp_re = re.compile(r'\[\d\d:\d\d:\d\d.\d\d\d --> \d\d:\d\d:\d\d.\d\d\d\]')

# Helper functions

def millisec(timeStr: str) -> int:
    """ Convert 'hh:mm:ss' string into milliseconds """
    try:
        h, m, s = timeStr.split(':')
        return (int(h) * 3600 + int(m) * 60 + int(s)) * 1000 # https://stackoverflow.com/a/6402859
    except:
        raise Exception(t('err_invalid_time_string', time = timeStr))

def ms_to_str(milliseconds: float, include_ms=False):
    """ Convert milliseconds into formatted timestamp 'hh:mm:ss' """
    seconds, milliseconds = divmod(milliseconds,1000)
    minutes, seconds = divmod(seconds, 60)
    hours, minutes = divmod(minutes, 60)
    formatted = f'{hours:02d}:{minutes:02d}:{seconds:02d}'
    if include_ms:
        formatted += f'.{milliseconds:03d}'
    return formatted 

def iter_except(function, exception):
        # Works like builtin 2-argument `iter()`, but stops on `exception`.
        try:
            while True:
                yield function()
        except exception:
            return
        
# Helper for text only output
        
def html_node_to_text(node: AdvancedHTMLParser.AdvancedTag) -> str:
    """
    Recursively get all text from a html node and its children. 
    """
    # For text nodes, return their value directly
    if AdvancedHTMLParser.isTextNode(node): # node.nodeType == node.TEXT_NODE:
        return html.unescape(node)
    # For element nodes, recursively process their children
    elif AdvancedHTMLParser.isTagNode(node):
        text_parts = []
        for child in node.childBlocks:
            text = html_node_to_text(child)
            if text:
                text_parts.append(text)
        # For block-level elements, prepend and append newlines
        if node.tagName.lower() in ['p', 'div', 'ul', 'ol', 'li', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'br']:
            if node.tagName.lower() == 'br':
                return '\n'
            else:
                return '\n' + ''.join(text_parts).strip() + '\n'
        else:
            return ''.join(text_parts)
    else:
        return ''

def html_to_text(parser: AdvancedHTMLParser.AdvancedHTMLParser) -> str:
    return html_node_to_text(parser.body)

# Helper for WebVTT output

def vtt_escape(txt: str) -> str:
    txt = html.escape(txt)
    while txt.find('\n\n') > -1:
        txt = txt.replace('\n\n', '\n')
    return txt    

def ms_to_webvtt(milliseconds) -> str:
    """converts milliseconds to the time stamp of WebVTT (HH:MM:SS.mmm)
    """
    # 1 hour = 3600000 milliseconds
    # 1 minute = 60000 milliseconds
    # 1 second = 1000 milliseconds
    hours, milliseconds = divmod(milliseconds, 3600000)
    minutes, milliseconds = divmod(milliseconds, 60000)
    seconds, milliseconds = divmod(milliseconds, 1000)
    return "{:02d}:{:02d}:{:02d}.{:03d}".format(hours, minutes, seconds, milliseconds)

def html_to_webvtt(parser: AdvancedHTMLParser.AdvancedHTMLParser, media_path: str):
    vtt = 'WEBVTT '
    paragraphs = parser.getElementsByTagName('p')
    # The first paragraph contains the title
    vtt += vtt_escape(paragraphs[0].textContent) + '\n\n'
    # Next paragraph contains info about the transcript. Add as a note.
    vtt += vtt_escape('NOTE\n' + html_node_to_text(paragraphs[1])) + '\n\n'
    # Add media source:
    vtt += f'NOTE media: {media_path}\n\n'

    #Add all segments as VTT cues
    segments = parser.getElementsByTagName('a')
    i = 0
    for i in range(len(segments)):
        segment = segments[i]
        name = segment.attributes['name']
        if name is not None:
            name_elems = name.split('_', 4)
            if len(name_elems) > 1 and name_elems[0] == 'ts':
                start = ms_to_webvtt(int(name_elems[1]))
                end = ms_to_webvtt(int(name_elems[2]))
                spkr = name_elems[3]
                txt = vtt_escape(html_node_to_text(segment))
                vtt += f'{i+1}\n{start} --> {end}\n<v {spkr}>{txt.lstrip()}\n\n'
    return vtt
    
class TimeEntry(ctk.CTkEntry): # special Entry box to enter time in the format hh:mm:ss
                               # based on https://stackoverflow.com/questions/63622880/how-to-make-python-automatically-put-colon-in-the-format-of-time-hhmmss
    def __init__(self, master, **kwargs):
        ctk.CTkEntry.__init__(self, master, **kwargs)
        vcmd = self.register(self.validate)

        self.bind('<Key>', self.format)
        self.configure(validate="all", validatecommand=(vcmd, '%P'))

        self.valid = re.compile(r'^\d{0,2}(:\d{0,2}(:\d{0,2})?)?$', re.I)

    def validate(self, text):
        if text == '':
            return True
        elif ''.join(text.split(':')).isnumeric():
            return not self.valid.match(text) is None
        else:
            return False

    def format(self, event):
        if event.keysym not in ['BackSpace', 'Shift_L', 'Shift_R', 'Control_L', 'Control_R']:
            i = self.index('insert')
            if i in [2, 5]:
                if event.char != ':':
                    if self.get()[i:i+1] != ':':
                        self.insert(i, ':')

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        self.user_models_dir = os.path.join(config_dir, 'whisper_models')
        os.makedirs(self.user_models_dir, exist_ok=True)
        whisper_models_readme = os.path.join(self.user_models_dir, 'readme.txt')
        if not os.path.exists(whisper_models_readme):
            with open(whisper_models_readme, 'w') as file:
                file.write('You can download custom Whisper-models for the transcription into this folder. \n' 
                           'See here for more information: https://github.com/kaixxx/noScribe/wiki/Add-custom-Whisper-models-for-transcription')            
        
        self.audio_file = ''
        self.audio_files = []  # List of dictionaries: [{'file': path, 'speaker': name}, ...]
        self.input_mode = 'single'  # 'single' or 'multi'
        self.transcript_file = ''
        self.log_file = None
        self.cancel = False # if set to True, transcription will be canceled

        # configure window
        self.title('noScribe - ' + t('app_header'))
        if platform.system() in ("Darwin", "Linux"):
            self.geometry(f"{1100}x{765}")
        else:
            self.geometry(f"{1100}x{690}")
        if platform.system() in ("Darwin", "Windows"):
            self.iconbitmap(os.path.join(app_dir, 'noScribeLogo.ico'))
        if platform.system() == "Linux":
            if hasattr(sys, "_MEIPASS"):
                self.iconphoto(True, tk.PhotoImage(file=os.path.join(sys._MEIPASS, "noScribeLogo.png")))
            else:
                self.iconphoto(True, tk.PhotoImage(file='noScribeLogo.png'))

        # header
        self.frame_header = ctk.CTkFrame(self, height=100)
        self.frame_header.pack(padx=0, pady=0, anchor='nw', fill='x')

        self.frame_header_logo = ctk.CTkFrame(self.frame_header, fg_color='transparent')
        self.frame_header_logo.pack(anchor='w', side='left')

        # logo
        self.logo_label = ctk.CTkLabel(self.frame_header_logo, text="noScribe", font=ctk.CTkFont(size=42, weight="bold"))
        self.logo_label.pack(padx=20, pady=[40, 0], anchor='w')

        # sub header
        self.header_label = ctk.CTkLabel(self.frame_header_logo, text=t('app_header'), font=ctk.CTkFont(size=16, weight="bold"))
        self.header_label.pack(padx=20, pady=[0, 20], anchor='w')

        # graphic
        self.header_graphic = ctk.CTkImage(dark_image=Image.open(os.path.join(app_dir, 'graphic_sw.png')), size=(926,119))
        self.header_graphic_label = ctk.CTkLabel(self.frame_header, image=self.header_graphic, text='')
        self.header_graphic_label.pack(anchor='ne', side='right', padx=[30,30])

        # main window
        self.frame_main = ctk.CTkFrame(self)
        self.frame_main.pack(padx=0, pady=0, anchor='nw', expand=True, fill='both')

        # create sidebar frame for options
        self.sidebar_frame = ctk.CTkFrame(self.frame_main, width=300, corner_radius=0, fg_color='transparent')
        self.sidebar_frame.pack(padx=0, pady=0, fill='y', expand=False, side='left')

        # create options scrollable frame
        self.scrollable_options = ctk.CTkScrollableFrame(self.sidebar_frame, width=300, corner_radius=0, fg_color='transparent')
        self.scrollable_options.pack(padx=0, pady=0, anchor='w', fill='both', expand=True)
        self.bind('<Configure>', self.on_resize) # Bind the configure event of options_frame to a check_scrollbar requirement function
        
        # Input Mode Selection
        self.label_input_mode = ctk.CTkLabel(self.scrollable_options, text='Input Mode:')
        self.label_input_mode.pack(padx=20, pady=[20,0], anchor='w')

        self.frame_input_mode = ctk.CTkFrame(self.scrollable_options, width=260, height=33, corner_radius=8, border_width=2)
        self.frame_input_mode.pack(padx=20, pady=[0,10], anchor='w')

        self.option_menu_input_mode = ctk.CTkOptionMenu(self.frame_input_mode, width=260, 
                                                       values=['Single File', 'Multi-File (by Speaker)'], 
                                                       command=self.on_input_mode_changed)
        self.option_menu_input_mode.pack(padx=5, pady=5)
        self.option_menu_input_mode.set('Single File')
        
        # input audio file
        self.label_audio_file = ctk.CTkLabel(self.scrollable_options, text=t('label_audio_file'))
        self.label_audio_file.pack(padx=20, pady=[20,0], anchor='w')

        self.frame_audio_file = ctk.CTkFrame(self.scrollable_options, width=260, height=33, corner_radius=8, border_width=2)
        self.frame_audio_file.pack(padx=20, pady=[0,10], anchor='w')

        self.button_audio_file_name = ctk.CTkButton(self.frame_audio_file, width=200, corner_radius=8, bg_color='transparent', 
                                                    fg_color='transparent', hover_color=self.frame_audio_file._bg_color, 
                                                    border_width=0, anchor='w',  
                                                    text=t('label_audio_file_name'), command=self.button_audio_file_event)
        self.button_audio_file_name.place(x=3, y=3)

        self.button_audio_file = ctk.CTkButton(self.frame_audio_file, width=45, height=29, text='📂', command=self.button_audio_file_event)
        self.button_audio_file.place(x=213, y=2)

        # Multi-File Input (initially hidden)
        self.label_audio_files = ctk.CTkLabel(self.scrollable_options, text='Audio files (by speaker):')
        self.label_audio_files.pack(padx=20, pady=[10,0], anchor='w')
        self.label_audio_files.pack_forget()  # Initially hidden

        # Frame for multi-file list
        self.frame_audio_files = ctk.CTkFrame(self.scrollable_options, width=260, corner_radius=8, border_width=2)
        self.frame_audio_files.pack(padx=20, pady=[0,10], anchor='w')
        self.frame_audio_files.pack_forget()  # Initially hidden

        # Scrollable frame for file list
        self.scrollable_files = ctk.CTkScrollableFrame(self.frame_audio_files, width=250, height=150)
        self.scrollable_files.pack(padx=5, pady=5, fill='both', expand=True)

        # Add file button
        self.button_add_audio_file = ctk.CTkButton(self.frame_audio_files, width=120, height=30, 
                                                  text='Add Audio File', command=self.button_add_audio_file_event)
        self.button_add_audio_file.pack(padx=5, pady=5, side='left')

        # input transcript file name
        self.label_transcript_file = ctk.CTkLabel(self.scrollable_options, text=t('label_transcript_file'))
        self.label_transcript_file.pack(padx=20, pady=[10,0], anchor='w')

        self.frame_transcript_file = ctk.CTkFrame(self.scrollable_options, width=260, height=33, corner_radius=8, border_width=2)
        self.frame_transcript_file.pack(padx=20, pady=[0,10], anchor='w')

        self.button_transcript_file_name = ctk.CTkButton(self.frame_transcript_file, width=200, corner_radius=8, bg_color='transparent', 
                                                    fg_color='transparent', hover_color=self.frame_transcript_file._bg_color, 
                                                    border_width=0, anchor='w',  
                                                    text=t('label_transcript_file_name'), command=self.button_transcript_file_event)
        self.button_transcript_file_name.place(x=3, y=3)

        self.button_transcript_file = ctk.CTkButton(self.frame_transcript_file, width=45, height=29, text='📂', command=self.button_transcript_file_event)
        self.button_transcript_file.place(x=213, y=2)

        # Options grid
        self.frame_options = ctk.CTkFrame(self.scrollable_options, width=250, fg_color='transparent')
        self.frame_options.pack_propagate(False)
        self.frame_options.pack(padx=20, pady=10, anchor='w', fill='x')

        # self.frame_options.grid_configure .resizable(width=False, height=True)
        self.frame_options.grid_columnconfigure(0, weight=1, minsize=0)
        self.frame_options.grid_columnconfigure(1, weight=0)

        # Start/stop
        self.label_start = ctk.CTkLabel(self.frame_options, text=t('label_start'))
        self.label_start.grid(column=0, row=0, sticky='w', pady=[0,5])

        self.entry_start = TimeEntry(self.frame_options, width=100)
        self.entry_start.grid(column='1', row='0', sticky='e', pady=[0,5])
        self.entry_start.insert(0, '00:00:00')

        self.label_stop = ctk.CTkLabel(self.frame_options, text=t('label_stop'))
        self.label_stop.grid(column=0, row=1, sticky='w', pady=[5,10])

        self.entry_stop = TimeEntry(self.frame_options, width=100)
        self.entry_stop.grid(column='1', row='1', sticky='e', pady=[5,10])

        # language
        self.label_language = ctk.CTkLabel(self.frame_options, text=t('label_language'))
        self.label_language.grid(column=0, row=2, sticky='w', pady=5)

        self.option_menu_language = ctk.CTkOptionMenu(self.frame_options, width=100, values=list(languages.keys()), dynamic_resizing=False)
        self.option_menu_language.grid(column=1, row=2, sticky='e', pady=5)
        last_language = get_config('last_language', 'auto')
        if last_language in languages.keys():
            self.option_menu_language.set(last_language)
        else:
            self.option_menu_language.set('Auto')
        
        # Whisper Model Selection   
        class CustomCTkOptionMenu(ctk.CTkOptionMenu):
            # Custom version that reads available models on drop down
            def __init__(self, noScribe_parent, master, width = 140, height = 28, corner_radius = None, bg_color = "transparent", fg_color = None, button_color = None, button_hover_color = None, text_color = None, text_color_disabled = None, dropdown_fg_color = None, dropdown_hover_color = None, dropdown_text_color = None, font = None, dropdown_font = None, values = None, variable = None, state = tk.NORMAL, hover = True, command = None, dynamic_resizing = True, anchor = "w", **kwargs):
                super().__init__(master, width, height, corner_radius, bg_color, fg_color, button_color, button_hover_color, text_color, text_color_disabled, dropdown_fg_color, dropdown_hover_color, dropdown_text_color, font, dropdown_font, values, variable, state, hover, command, dynamic_resizing, anchor, **kwargs)
                self.noScribe_parent = noScribe_parent
                self.old_value = ''

            def _clicked(self, event=0):
                self.old_value = self.get()
                self._values = self.noScribe_parent.get_whisper_models()
                self._values.append('--------------------')
                self._values.append(t('label_add_custom_models'))
                self._dropdown_menu.configure(values=self._values)
                super()._clicked(event)
                
            def _dropdown_callback(self, value: str):
                if value == self._values[-2]:  # divider
                    return
                if value == self._values[-1]:  # Add custom model
                    # show custom model folder
                    path = self.noScribe_parent.user_models_dir
                    try:
                        os_type = platform.system()
                        if os_type == "Windows":
                            os.startfile(path)
                        elif os_type == "Darwin":
                            run(["open", path])
                        elif os_type == "Linux":
                            run(["xdg-open", path])
                        else:
                            raise OSError(f"Unsupported operating system: {os_type}")
                    except Exception as e:
                        self.noScribe_parent.logn(f"Failed to open folder: {e}")
                else:
                    super()._dropdown_callback(value)
        
        self.label_whisper_model = ctk.CTkLabel(self.frame_options, text=t('label_whisper_model'))
        self.label_whisper_model.grid(column=0, row=3, sticky='w', pady=5)

        models = self.get_whisper_models()
        self.option_menu_whisper_model = CustomCTkOptionMenu(self, 
                                                       self.frame_options, 
                                                       width=100,
                                                       values=models,
                                                       dynamic_resizing=False)
        self.option_menu_whisper_model.grid(column=1, row=3, sticky='e', pady=5)
        last_whisper_model = get_config('last_whisper_model', 'precise')
        if last_whisper_model in models:
            self.option_menu_whisper_model.set(last_whisper_model)
        elif len(models) > 0:
            self.option_menu_whisper_model.set(models[0])

        # Mark pauses
        self.label_pause = ctk.CTkLabel(self.frame_options, text=t('label_pause'))
        self.label_pause.grid(column=0, row=4, sticky='w', pady=5)

        self.option_menu_pause = ctk.CTkOptionMenu(self.frame_options, width=100, values=['none', '1sec+', '2sec+', '3sec+'])
        self.option_menu_pause.grid(column=1, row=4, sticky='e', pady=5)
        self.option_menu_pause.set(get_config('last_pause', '1sec+'))

        # Speaker Detection (Diarization)
        self.label_speaker = ctk.CTkLabel(self.frame_options, text=t('label_speaker'))
        self.label_speaker.grid(column=0, row=5, sticky='w', pady=5)

        self.option_menu_speaker = ctk.CTkOptionMenu(self.frame_options, width=100, values=['none', 'auto', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10'])
        self.option_menu_speaker.grid(column=1, row=5, sticky='e', pady=5)
        self.option_menu_speaker.set(get_config('last_speaker', 'auto'))

        # Overlapping Speech (Diarization)
        self.label_overlapping = ctk.CTkLabel(self.frame_options, text=t('label_overlapping'))
        self.label_overlapping.grid(column=0, row=6, sticky='w', pady=5)

        self.check_box_overlapping = ctk.CTkCheckBox(self.frame_options, text = '')
        self.check_box_overlapping.grid(column=1, row=6, sticky='e', pady=5)
        overlapping = config.get('last_overlapping', True)
        if overlapping:
            self.check_box_overlapping.select()
        else:
            self.check_box_overlapping.deselect()
            
        # Disfluencies
        self.label_disfluencies = ctk.CTkLabel(self.frame_options, text=t('label_disfluencies'))
        self.label_disfluencies.grid(column=0, row=7, sticky='w', pady=5)

        self.check_box_disfluencies = ctk.CTkCheckBox(self.frame_options, text = '')
        self.check_box_disfluencies.grid(column=1, row=7, sticky='e', pady=5)
        check_box_disfluencies = config.get('last_disfluencies', True)
        if check_box_disfluencies:
            self.check_box_disfluencies.select()
        else:
            self.check_box_disfluencies.deselect()

        # Timestamps in text
        self.label_timestamps = ctk.CTkLabel(self.frame_options, text=t('label_timestamps'))
        self.label_timestamps.grid(column=0, row=8, sticky='w', pady=5)

        self.check_box_timestamps = ctk.CTkCheckBox(self.frame_options, text = '')
        self.check_box_timestamps.grid(column=1, row=8, sticky='e', pady=5)
        check_box_timestamps = config.get('last_timestamps', False)
        if check_box_timestamps:
            self.check_box_timestamps.select()
        else:
            self.check_box_timestamps.deselect()
        
        # Start Button
        self.start_button = ctk.CTkButton(self.sidebar_frame, height=42, text=t('start_button'), command=self.button_start_event)
        self.start_button.pack(padx=[20, 0], pady=[20,30], expand=False, fill='x', anchor='sw')

        # Stop Button
        self.stop_button = ctk.CTkButton(self.sidebar_frame, height=42, fg_color='darkred', hover_color='darkred', text=t('stop_button'), command=self.button_stop_event)
        
        # create log textbox
        self.log_frame = ctk.CTkFrame(self.frame_main, corner_radius=0, fg_color='transparent')
        self.log_frame.pack(padx=0, pady=0, fill='both', expand=True, side='top')

        self.log_textbox = ctk.CTkTextbox(self.log_frame, wrap='word', state="disabled", font=("",16), text_color="lightgray")
        self.log_textbox.tag_config('highlight', foreground='darkorange')
        self.log_textbox.tag_config('error', foreground='yellow')
        self.log_textbox.pack(padx=20, pady=[20,0], expand=True, fill='both')

        self.hyperlink = HyperlinkManager(self.log_textbox._textbox)

        # Frame progress bar / edit button
        self.frame_edit = ctk.CTkFrame(self.frame_main, height=20, corner_radius=0, fg_color=self.log_textbox._fg_color)
        self.frame_edit.pack(padx=20, pady=[0,30], anchor='sw', fill='x', side='bottom')

        # Edit Button
        self.edit_button = ctk.CTkButton(self.frame_edit, fg_color=self.log_textbox._scrollbar_button_color, 
                                         text=t('editor_button'), command=self.launch_editor, width=140)
        self.edit_button.pack(padx=[20,10], pady=[10,10], expand=False, anchor='se', side='right')

        # Progress bar
        self.progress_textbox = ctk.CTkTextbox(self.frame_edit, wrap='none', height=15, state="disabled", font=("",16), text_color="lightgray")
        self.progress_textbox.pack(padx=[10,10], pady=[5,0], expand=True, fill='x', anchor='sw', side='left')

        self.update_scrollbar_visibility()        
        #self.progress_bar = ctk.CTkProgressBar(self.frame_edit, mode='determinate', progress_color='darkred', fg_color=self.log_textbox._fg_color)
        #self.progress_bar.set(0)
        # self.progress_bar.pack(padx=[10,10], pady=[10,10], expand=True, fill='x', anchor='sw', side='left')

        # status bar bottom
        #self.frame_status = ctk.CTkFrame(self, height=20, corner_radius=0)
        #self.frame_status.pack(padx=0, pady=[0,0], anchor='sw', fill='x', side='bottom')

        self.logn(t('welcome_message'), 'highlight')
        self.log(t('welcome_credits', v=app_version, y=app_year))
        self.logn('https://github.com/kaixxx/noScribe', link='https://github.com/kaixxx/noScribe#readme')
        self.logn(t('welcome_instructions'))
        
        # check for new releases
        if get_config('check_for_update', 'True') == 'True':
            try:
                latest_release = json.loads(urllib.request.urlopen(
                    urllib.request.Request('https://api.github.com/repos/kaixxx/noScribe/releases/latest',
                    headers={'Accept': 'application/vnd.github.v3+json'},),
                    timeout=2).read())
                latest_release_version = str(latest_release['tag_name']).lstrip('v')
                if version_higher(latest_release_version, app_version) == 1:
                    self.logn(t('new_release', v=latest_release_version), 'highlight')
                    self.logn(str(latest_release['body'])) # release info
                    self.log(t('new_release_download'))
                    self.logn(str(latest_release['html_url']), link=str(latest_release['html_url']))
                    self.logn()
            except:
                pass
            
    # Events and Methods

    def get_whisper_models(self):
        self.whisper_model_paths = {}
        
        def collect_models(dir):        
            for entry in os.listdir(dir):
                entry_path = os.path.join(dir, entry)
                if os.path.isdir(entry_path):
                    if entry in self.whisper_model_paths:
                        self.logn(f'Ignored double name for whisper model: "{entry}"', 'error')
                    else:
                        self.whisper_model_paths[entry]=entry_path 
       
        # collect system models:
        collect_models(os.path.join(app_dir, 'models'))
        
        # collect user defined models:        
        collect_models(self.user_models_dir)

        return list(self.whisper_model_paths.keys())
    
    def on_whisper_model_selected(self, value):
        print(self.option_menu_whisper_model.old_value)
        print(value)
        
    def on_resize(self, event):
        self.update_scrollbar_visibility()

    def update_scrollbar_visibility(self):
        # Get the size of the scroll region and current canvas size
        canvas = self.scrollable_options._parent_canvas  
        scroll_region_height = canvas.bbox("all")[3]
        canvas_height = canvas.winfo_height()        
        
        scrollbar = self.scrollable_options._scrollbar

        if scroll_region_height > canvas_height:
            scrollbar.grid()
        else:
            scrollbar.grid_remove()  # Hide the scrollbar if not needed    

    def launch_editor(self, file=''):
        # Launch the editor in a seperate process so that in can stay running even if noScribe quits.
        # Source: https://stackoverflow.com/questions/13243807/popen-waiting-for-child-process-even-when-the-immediate-child-has-terminated/13256908#13256908 
        # set system/version dependent "start_new_session" analogs
        if file == '':
            file = self.transcript_file
        ext = os.path.splitext(self.transcript_file)[1][1:]
        if file != '' and ext != 'html':
            file = ''
            if not tk.messagebox.askyesno(title='noScribe', message=t('err_editor_invalid_format')):
                return
        program: str = None
        if platform.system() == 'Windows':
            program = os.path.join(app_dir, 'noScribeEdit', 'noScribeEdit.exe')
        elif platform.system() == "Darwin": # = MAC
            # use local copy in development, installed one if used as an app:
            program = os.path.join(app_dir, 'noScribeEdit', 'noScribeEdit')
            if not os.path.exists(program):
                program = os.path.join(os.sep, 'Applications', 'noScribeEdit.app', 'Contents', 'MacOS', 'noScribeEdit')
        elif platform.system() == "Linux":
            if hasattr(sys, "_MEIPASS"):
                program = os.path.join(sys._MEIPASS, 'noScribeEdit', "noScribeEdit")
            else:
                program = os.path.join(app_dir, 'noScribeEdit', "noScribeEdit.py")
        kwargs = {}
        if platform.system() == 'Windows':
            # from msdn [1]
            CREATE_NEW_PROCESS_GROUP = 0x00000200  # note: could get it from subprocess
            DETACHED_PROCESS = 0x00000008          # 0x8 | 0x200 == 0x208
            kwargs.update(creationflags=DETACHED_PROCESS | CREATE_NEW_PROCESS_GROUP)  
        else:  # should work on all POSIX systems, Linux and macOS 
            kwargs.update(start_new_session=True)

        if program is not None and os.path.exists(program):
            popenargs = [program]
            if platform.system() == "Linux" and not hasattr(sys, "_MEIPASS"): # only do this, if you run as python script; Linux python vs. executable needs refinement
                popenargs = [sys.executable, program]
            if file != '':
                popenargs.append(file)
            Popen(popenargs, **kwargs)
        else:
            self.logn(t('err_noScribeEdit_not_found'), 'error')

    def openLink(self, link: str) -> None:
        if link.startswith('file://') and link.endswith('.html'):
            self.launch_editor(link[7:])
        else: 
            webbrowser.open(link)
    


    def log(self, txt: str = '', tags: list = [], where: str = 'both', link: str = '', tb: str = '') -> None:
        """ Log to main window (where can be 'screen', 'file', or 'both') 
        tb = formatted traceback of the error, only logged to file
        """
        
        # Handle screen logging if requested and textbox exists
        if where != 'file' and hasattr(self, 'log_textbox') and self.log_textbox.winfo_exists():
            try:
                self.log_textbox.configure(state=tk.NORMAL)
                
                if link:
                    tags = tags + self.hyperlink.add(partial(self.openLink, link))
                
                self.log_textbox.insert(tk.END, txt, tags)
                self.log_textbox.yview_moveto(1)  # Scroll to last line
                
                # Schedule disabling the textbox in the main thread
                self.log_textbox.after(0, lambda: self.log_textbox.configure(state=tk.DISABLED))
            except Exception as e:
                # Log screen errors only to file to prevent recursion
                if where == 'both':
                    self.log(f"Error updating log_textbox: {str(e)}\nOriginal error: {txt}", tags='error', where='file', tb=tb)

        # Handle file logging if requested
        if where != 'screen' and self.log_file and not self.log_file.closed:
            try:
                if tags == 'error':
                    txt = f'ERROR: {txt}'
                if tb != '':
                    txt = f'{txt}\nTraceback:\n{tb}' 
                self.log_file.write(txt)
                self.log_file.flush()
            except Exception as e:
                # If we get here, both screen and file logging failed
                # As a last resort, print to stderr to not lose the error
                import sys
                print(f"Critical error - both screen and file logging failed: {str(e)}\nOriginal error: {txt}\nOriginal traceback:\n{tb}", file=sys.stderr)

    def logn(self, txt: str = '', tags: list = [], where: str = 'both', link:str = '', tb: str = '') -> None:
        """ Log with a newline appended """
        self.log(f'{txt}\n', tags, where, link, tb)

    def logr(self, txt: str = '', tags: list = [], where: str = 'both', link:str = '', tb: str = '') -> None:
        """ Replace the last line of the log """
        if where != 'file':
            self.log_textbox.configure(state=ctk.NORMAL)
            self.log_textbox.delete("end-1c linestart", "end-1c")
        self.log(txt, tags, where, link, tb)

    def button_audio_file_event(self):
        fn = tk.filedialog.askopenfilename(initialdir=os.path.dirname(self.audio_file), initialfile=os.path.basename(self.audio_file))
        if fn:
            self.audio_file = fn
            self.logn(t('log_audio_file_selected') + self.audio_file)
            self.button_audio_file_name.configure(text=os.path.basename(self.audio_file))

    def on_input_mode_changed(self, value):
        """Handle input mode change between single and multi-file"""
        if value == 'Single File':
            self.input_mode = 'single'
            self.label_audio_file.pack(padx=20, pady=[10,0], anchor='w')
            self.frame_audio_file.pack(padx=20, pady=[0,10], anchor='w')
            self.label_audio_files.pack_forget()
            self.frame_audio_files.pack_forget()
        else:
            self.input_mode = 'multi'
            self.label_audio_file.pack_forget()
            self.frame_audio_file.pack_forget()
            self.label_audio_files.pack(padx=20, pady=[10,0], anchor='w')
            self.frame_audio_files.pack(padx=20, pady=[0,10], anchor='w')

    def button_add_audio_file_event(self):
        """Add a new audio file to the multi-file list"""
        fn = tk.filedialog.askopenfilename(
            title='Add Audio File',
            filetypes=[
                ('Audio files', '*.wav *.mp3 *.m4a *.aac *.flac *.ogg *.wma'),
                ('All files', '*.*')
            ]
        )
        if fn:
            # Generate speaker name from filename
            speaker_name = Path(fn).stem
            if speaker_name.startswith('audio') and len(speaker_name) > 5:
                # Handle Zoom-style naming: audio1Speaker1.m4a -> Speaker1
                if 'Speaker' in speaker_name:
                    speaker_name = speaker_name.split('Speaker')[-1]
                    speaker_name = f"Speaker{speaker_name}"
                else:
                    speaker_name = f"Speaker{len(self.audio_files) + 1:02d}"
            else:
                speaker_name = f"Speaker{len(self.audio_files) + 1:02d}"
            
            # Add to list
            self.audio_files.append({'file': fn, 'speaker': speaker_name})
            self.update_audio_files_display()
            self.logn(f"Audio files selected: {len(self.audio_files)} files")

    def update_audio_files_display(self):
        """Update the display of audio files in the multi-file UI"""
        # Clear existing widgets
        for widget in self.scrollable_files.winfo_children():
            widget.destroy()
        
        # Add file entries
        for i, audio_info in enumerate(self.audio_files):
            frame = ctk.CTkFrame(self.scrollable_files, width=240, height=60)
            frame.pack(padx=5, pady=2, fill='x')
            
            # Speaker name entry
            speaker_entry = ctk.CTkEntry(frame, width=100, height=25, placeholder_text='Speaker Name')
            speaker_entry.pack(side='left', padx=5, pady=5)
            speaker_entry.insert(0, audio_info['speaker'])
            speaker_entry.bind('<FocusOut>', lambda e, idx=i: self.update_speaker_name(idx, e.widget.get()))
            
            # File path (truncated)
            file_label = ctk.CTkLabel(frame, text=os.path.basename(audio_info['file']), width=100)
            file_label.pack(side='left', padx=5, pady=5)
            
            # Remove button
            remove_btn = ctk.CTkButton(frame, width=30, height=25, text='×', 
                                     command=lambda idx=i: self.remove_audio_file(idx))
            remove_btn.pack(side='right', padx=5, pady=5)

    def update_speaker_name(self, index, new_name):
        """Update speaker name for a file"""
        if 0 <= index < len(self.audio_files):
            self.audio_files[index]['speaker'] = new_name

    def remove_audio_file(self, index):
        """Remove an audio file from the list"""
        if 0 <= index < len(self.audio_files):
            del self.audio_files[index]
            self.update_audio_files_display()

    def button_transcript_file_event(self):
        if self.transcript_file != '':
            _initialdir = os.path.dirname(self.transcript_file)
            _initialfile = os.path.basename(self.transcript_file)
        else:
            # Determine initial directory - avoid read-only locations like /Volumes/
            if self.input_mode == 'single' and self.audio_file:
                audio_dir = os.path.dirname(self.audio_file)
                # Check if audio file is in a read-only location (like mounted DMG)
                if audio_dir.startswith('/Volumes/'):
                    _initialdir = str(Path.home() / 'Desktop')  # Use Desktop instead
                    self.logn("Audio file is in read-only location. Using Desktop for transcript.", 'error')
                else:
                    _initialdir = audio_dir
                _initialfile = Path(os.path.basename(self.audio_file)).stem
            elif self.input_mode == 'multi' and self.audio_files:
                # For multi-file mode, use Desktop as default to avoid issues
                _initialdir = str(Path.home() / 'Desktop')
                _initialfile = 'multi_speaker_transcript'
            else:
                # No audio files selected yet, use Desktop
                _initialdir = str(Path.home() / 'Desktop')
                _initialfile = 'transcript'
        if not ('last_filetype' in config):
            config['last_filetype'] = 'html'
        filetypes = [
            ('noScribe Transcript','*.html'), 
            ('Text only','*.txt'),
            ('WebVTT Subtitles (also for EXMARaLDA)', '*.vtt')
        ]
        for i, ft in enumerate(filetypes):
            if ft[1] == f'*.{config["last_filetype"]}':
                filetypes.insert(0, filetypes.pop(i))
                break
        fn = tk.filedialog.asksaveasfilename(initialdir=_initialdir, initialfile=_initialfile, 
                                             filetypes=filetypes, 
                                             defaultextension=config['last_filetype'])
        if fn:
            self.transcript_file = fn
            self.logn(t('log_transcript_filename') + self.transcript_file)
            self.button_transcript_file_name.configure(text=os.path.basename(self.transcript_file))
            config['last_filetype'] = os.path.splitext(self.transcript_file)[1][1:]
            
    def set_progress(self, step, value):
        """ Update state of the progress bar """
        progr = -1
        if step == 1:
            progr = value * 0.05 / 100
        elif step == 2:
            progr = 0.05 # (step 1)
            progr = progr + (value * 0.45 / 100)
        elif step == 3:
            if self.speaker_detection != 'none':
                progr = 0.05 + 0.45 # (step 1 + step 2)
                progr_factor = 0.5
            else:
                progr = 0.05 # (step 1)
                progr_factor = 0.95
            progr = progr + (value * progr_factor / 100)
        if progr >= 1:
            progr = 0.99 # whisper sometimes still needs some time to finish even if the progress is already at 100%. This can be confusing, so we never go above 99%...

        # Update progress_textbox
        if progr < 0:
            progr_str = ''
        else:
            progr_str = f'({t("overall_progress")}{round(progr * 100)}%)'
        self.progress_textbox.configure(state=ctk.NORMAL)        
        self.progress_textbox.delete('1.0', tk.END)
        self.progress_textbox.insert(tk.END, progr_str)
        self.progress_textbox.configure(state=ctk.DISABLED)


    ################################################################################################
    # Main function

    def transcribe_multi_files(self, tmpdir):
        """Transcribe multiple files using word-level timestamps for better timeline accuracy"""
        from faster_whisper import WhisperModel
        
        # Initialize Whisper model
        if platform.system() == "Darwin":
            whisper_device = 'auto'
        elif platform.system() in ('Windows', 'Linux'):
            whisper_device = self.whisper_xpu
        else:
            raise Exception('Platform not supported yet.')
        
        self.logn(f"Loading Whisper model: {self.whisper_model}")
        self.logn(f"Device: {whisper_device}, Compute type: {self.whisper_compute_type}")
        
        try:
            model = WhisperModel(self.whisper_model, device=whisper_device, 
                               cpu_threads=number_threads, compute_type=self.whisper_compute_type, 
                               local_files_only=True)
            self.logn("Whisper model loaded successfully")
        except Exception as e:
            self.logn(f"Error loading Whisper model: {str(e)}", 'error')
            raise
        
        all_word_segments = []
        
        # Determine language parameter
        whisper_lang = None
        if self.language_name != 'Auto' and self.language_name != 'Multilingual':
            whisper_lang = languages.get(self.language_name, None)
        
        self.logn(f"Language setting: {self.language_name} -> {whisper_lang}")
        
        # Transcribe each file separately - using WORD-LEVEL timestamps for precision
        for i, audio_info in enumerate(self.audio_files):
            if self.cancel:
                self.logn("Transcription canceled by user")
                break
                
            self.logn(f"Processing file {i+1}/{len(self.audio_files)}: {audio_info['speaker']}")
            self.logn(f"File: {os.path.basename(audio_info['file'])}")
            
            tmp_audio = os.path.join(tmpdir.name, f'tmp_audio_{i}.wav')
            
            try:
                # Convert audio file
                self.convert_audio_file(audio_info['file'], tmp_audio)
                self.logn(f"Audio conversion completed for {audio_info['speaker']}")
                
                # Transcribe with word-level timestamps for precise timeline management
                segments, info = model.transcribe(
                    tmp_audio,
                    language=whisper_lang,
                    word_timestamps=True,  # Critical for timeline accuracy
                    vad_filter=True,
                    vad_parameters={'min_silence_duration_ms': 1000, 'speech_pad_ms': 400},
                    beam_size=5,
                    temperature=0.0,
                    condition_on_previous_text=True,
                    no_speech_threshold=0.6
                )
                
                self.logn(f"Transcription completed for {audio_info['speaker']}")
                
                # Extract WORD-LEVEL segments for precise timeline control
                word_count = 0
                for segment in segments:
                    # Process each word individually with precise timestamps
                    if hasattr(segment, 'words') and segment.words:
                        for word in segment.words:
                            if word.word.strip():  # Only non-empty words
                                word_segment = {
                                    'start': word.start,
                                    'end': word.end,
                                    'text': word.word.strip(),
                                    'speaker': audio_info['speaker']
                                }
                                all_word_segments.append(word_segment)
                                word_count += 1
                    else:
                        # Fallback to segment-level if words not available
                        if segment.text.strip():
                            seg_dict = {
                                'start': segment.start,
                                'end': segment.end,
                                'text': segment.text.strip(),
                                'speaker': audio_info['speaker']
                            }
                            all_word_segments.append(seg_dict)
                            word_count += 1
                
                self.logn(f"Added {word_count} word-level segments from {audio_info['speaker']}")
                
            except Exception as e:
                self.logn(f"Error processing {audio_info['speaker']}: {str(e)}", 'error')
                # Continue with other files instead of failing completely
                continue
        
        # Sort ALL segments by start time - this is the key to proper timeline order
        all_word_segments.sort(key=lambda x: x['start'])
        self.logn(f"Total word segments collected: {len(all_word_segments)}")
        
        # Now merge word segments into speaker utterances for better readability
        merged_segments = self.merge_word_segments_to_utterances(all_word_segments)
        self.logn(f"Merged into {len(merged_segments)} speaker utterances")
        
        return merged_segments

    def convert_audio_file(self, input_file, output_file):
        """Convert audio file to WAV format"""
        import shlex
        from subprocess import Popen, PIPE, STDOUT
        if platform.system() == 'Windows':
            from subprocess import STARTUPINFO, STARTF_USESHOWWINDOW
        
        # Validate input file exists
        if not os.path.exists(input_file):
            raise Exception(f'Input audio file not found: {input_file}')
            
        arguments = f' -loglevel warning -hwaccel auto -y -i "{input_file}" -ar 16000 -ac 1 -c:a pcm_s16le "{output_file}"'
        
        if platform.system() == 'Windows':
            ffmpeg_path = os.path.join(app_dir, 'ffmpeg.exe')
            ffmpeg_cmd = ffmpeg_path + arguments
        elif platform.system() == "Darwin":
            # Try bundled ffmpeg first, fall back to system ffmpeg
            bundled_ffmpeg = os.path.join(app_dir, 'ffmpeg')
            
            # Check if system is ARM64 and bundled ffmpeg is x86_64 (incompatible)
            import subprocess
            try:
                arch_result = subprocess.run(['uname', '-m'], capture_output=True, text=True)
                is_arm64 = 'arm64' in arch_result.stdout
                
                if is_arm64 and os.path.exists(bundled_ffmpeg):
                    # Test if bundled ffmpeg works
                    test_result = subprocess.run([bundled_ffmpeg, '-version'], 
                                               capture_output=True, text=True, timeout=5)
                    if test_result.returncode != 0:
                        self.logn("Bundled FFmpeg incompatible with ARM64, using system FFmpeg")
                        raise Exception("Bundled FFmpeg not compatible")
            except:
                # If bundled ffmpeg fails, try system ffmpeg
                pass
            
            # Try system ffmpeg first on ARM64, or if bundled fails
            try:
                system_ffmpeg = subprocess.run(['which', 'ffmpeg'], capture_output=True, text=True)
                if system_ffmpeg.returncode == 0:
                    ffmpeg_path = system_ffmpeg.stdout.strip()
                    self.logn(f"Using system FFmpeg: {ffmpeg_path}")
                else:
                    ffmpeg_path = bundled_ffmpeg
            except:
                ffmpeg_path = bundled_ffmpeg
                
            ffmpeg_cmd = shlex.split(ffmpeg_path + arguments)
        elif platform.system() == "Linux":
            ffmpeg_path = os.path.join(app_dir, 'ffmpeg-linux-x86_64')
            ffmpeg_cmd = shlex.split(ffmpeg_path + arguments)
        else:
            raise Exception('Platform not supported yet.')
        
        # Check if ffmpeg exists and is executable
        if not os.path.exists(ffmpeg_path):
            raise Exception(f'FFmpeg not found at: {ffmpeg_path}')
        
        if not os.access(ffmpeg_path, os.X_OK):
            raise Exception(f'FFmpeg not executable: {ffmpeg_path}')
        
        self.logn(f'Converting {os.path.basename(input_file)} to WAV format')
        
        try:
            if platform.system() == 'Windows':
                startupinfo = STARTUPINFO()
                startupinfo.dwFlags |= STARTF_USESHOWWINDOW
                with Popen(ffmpeg_cmd, stdout=PIPE, stderr=STDOUT, bufsize=1, universal_newlines=True, encoding='utf-8', startupinfo=startupinfo) as proc:
                    output_lines = []
                    for line in proc.stdout:
                        output_lines.append(line.strip())
            else:
                with Popen(ffmpeg_cmd, stdout=PIPE, stderr=STDOUT, bufsize=1, universal_newlines=True, encoding='utf-8') as proc:
                    output_lines = []
                    for line in proc.stdout:
                        output_lines.append(line.strip())
            
            if proc.returncode != 0:
                error_output = '\n'.join(output_lines[-10:])  # Last 10 lines of output
                self.logn(f'FFmpeg error output: {error_output}', 'error')
                raise Exception(f'FFmpeg conversion failed with return code {proc.returncode}')
                
            # Verify output file was created
            if not os.path.exists(output_file):
                raise Exception('Output file was not created by FFmpeg')
                
        except Exception as e:
            self.logn(f'Error converting audio file {input_file}: {str(e)}', 'error')
            raise Exception(t('err_converting_audio'))

    def create_transcript_from_segments(self, segments):
        """Create transcript HTML from timeline-sorted segments with timestamp info"""
        # Prepare transcript html
        d = AdvancedHTMLParser.AdvancedHTMLParser()
        d.parseStr(default_html)
        
        # Add audio file path(s)
        tag = d.createElement("meta")
        tag.name = "audio_source"
        if self.input_mode == 'single':
            tag.content = self.audio_file
        else:
            tag.content = "; ".join([info['file'] for info in self.audio_files])
        d.head.appendChild(tag)

        # Add main body
        main_body = d.createElement('div')
        main_body.addClass('WordSection1')
        d.body.appendChild(main_body)

        # Header
        p = d.createElement('p')
        p.setStyle('font-weight', '600')
        if self.input_mode == 'single':
            p.appendText(Path(self.audio_file).stem)
        else:
            p.appendText("Multi-Speaker Transcript (Timeline Sorted)")
        main_body.appendChild(p)

        # Add timeline info
        if segments:
            timeline_info = d.createElement('p')
            timeline_info.setStyle('font-style', 'italic')
            timeline_info.setStyle('color', '#666')
            start_time = segments[0]['start']
            end_time = segments[-1]['end']
            timeline_info.appendText(f"Timeline: {start_time:.1f}s - {end_time:.1f}s ({len(segments)} utterances)")
            main_body.appendChild(timeline_info)

        # Process segments - they are already timeline-sorted and speaker-merged
        for segment in segments:
            segment_text = segment['text'].strip()
            if not segment_text:
                continue
            
            # Create paragraph for each utterance
            paragraph = d.createElement('p')
            
            # Add timestamp if enabled
            if self.check_box_timestamps and self.check_box_timestamps.get():
                timestamp_span = d.createElement('span')
                timestamp_span.setStyle('color', '#888')
                timestamp_span.setStyle('font-size', '0.8em')
                timestamp_span.appendText(f"[{segment['start']:.1f}s-{segment['end']:.1f}s] ")
                paragraph.appendChild(timestamp_span)
            
            # Add speaker label
            speaker_span = d.createElement('span')
            speaker_span.setStyle('font-weight', 'bold')
            speaker_span.setStyle('color', '#2E7D32')
            speaker_span.appendText(f"{segment['speaker']}: ")
            paragraph.appendChild(speaker_span)
            
            # Add text content
            text_span = d.createElement('span')
            text_span.appendText(segment_text)
            paragraph.appendChild(text_span)
            
            main_body.appendChild(paragraph)

        # Save transcript using the corrected method
        self.save_multi_file_transcript(d)

    def save_multi_file_transcript(self, d):
        """Save the multi-file transcript to file"""
        txt = ''
        if self.file_ext == 'html':
            txt = d.asHTML()
        elif self.file_ext == 'txt':
            txt = html_to_text(d)
        elif self.file_ext == 'vtt':
            txt = html_to_webvtt(d, self.audio_files[0]['file'] if self.audio_files else 'multi_file_transcript')
        else:
            raise TypeError(f'Invalid file type "{self.file_ext}".')
        
        try:
            if txt != '':
                with open(self.my_transcript_file, 'w', encoding="utf-8") as f:
                    f.write(txt)
                    f.flush()
                self.logn('Multi-file transcript saved successfully')
        except Exception as e:
            self.logn(f'Error saving transcript: {str(e)}', 'error')
            # Try alternative filename
            try:
                alt_filename = f"{Path(self.my_transcript_file).stem}_multifile.{self.file_ext}"
                alt_path = Path(self.my_transcript_file).parent / alt_filename
                with open(alt_path, 'w', encoding="utf-8") as f:
                    f.write(txt)
                    f.flush()
                self.logn(f'Transcript saved to alternative location: {alt_path}')
            except Exception as e2:
                self.logn(f'Failed to save transcript: {str(e2)}', 'error')

    def convert_audio_file_single_mode(self):
        """Convert audio file to WAV format for single file mode with start/stop times"""
        import shlex
        import subprocess
        from subprocess import Popen, PIPE, STDOUT
        if platform.system() == 'Windows':
            from subprocess import STARTUPINFO, STARTF_USESHOWWINDOW
        
        # Validate input file exists
        if not os.path.exists(self.audio_file):
            raise Exception(f'Input audio file not found: {self.audio_file}')
        
        # Build FFmpeg arguments with start/stop times
        if int(self.stop) > 0:  # transcribe only part of the audio
            end_pos_cmd = f'-to {self.stop}ms'
        else:  # transcribe until the end
            end_pos_cmd = ''
            
        arguments = f' -loglevel warning -hwaccel auto -y -ss {self.start}ms {end_pos_cmd} -i "{self.audio_file}" -ar 16000 -ac 1 -c:a pcm_s16le "{self.tmp_audio_file}"'
        
        if platform.system() == 'Windows':
            ffmpeg_path = os.path.join(app_dir, 'ffmpeg.exe')
            ffmpeg_cmd = ffmpeg_path + arguments
        elif platform.system() == "Darwin":
            # Try bundled ffmpeg first, fall back to system ffmpeg
            bundled_ffmpeg = os.path.join(app_dir, 'ffmpeg')
            
            # Check if system is ARM64 and bundled ffmpeg is x86_64 (incompatible)
            try:
                arch_result = subprocess.run(['uname', '-m'], capture_output=True, text=True)
                is_arm64 = 'arm64' in arch_result.stdout
                
                if is_arm64 and os.path.exists(bundled_ffmpeg):
                    # Test if bundled ffmpeg works
                    test_result = subprocess.run([bundled_ffmpeg, '-version'], 
                                               capture_output=True, text=True, timeout=5)
                    if test_result.returncode != 0:
                        self.logn("Bundled FFmpeg incompatible with ARM64, using system FFmpeg")
                        raise Exception("Bundled FFmpeg not compatible")
            except:
                # If bundled ffmpeg fails, try system ffmpeg
                pass
            
            # Try system ffmpeg first on ARM64, or if bundled fails
            try:
                system_ffmpeg = subprocess.run(['which', 'ffmpeg'], capture_output=True, text=True)
                if system_ffmpeg.returncode == 0:
                    ffmpeg_path = system_ffmpeg.stdout.strip()
                    self.logn(f"Using system FFmpeg: {ffmpeg_path}")
                else:
                    ffmpeg_path = bundled_ffmpeg
            except:
                ffmpeg_path = bundled_ffmpeg
                
            ffmpeg_cmd = shlex.split(ffmpeg_path + arguments)
        elif platform.system() == "Linux":
            ffmpeg_path = os.path.join(app_dir, 'ffmpeg-linux-x86_64')
            ffmpeg_cmd = shlex.split(ffmpeg_path + arguments)
        else:
            raise Exception('Platform not supported yet.')
        
        # Check if ffmpeg exists and is executable
        if not os.path.exists(ffmpeg_path):
            raise Exception(f'FFmpeg not found at: {ffmpeg_path}')
        
        if not os.access(ffmpeg_path, os.X_OK):
            raise Exception(f'FFmpeg not executable: {ffmpeg_path}')
        
        self.logn(f'Converting audio file: {os.path.basename(self.audio_file)}')
        self.logn(ffmpeg_cmd, where='file')
        
        try:
            if platform.system() == 'Windows':
                # (suppresses the terminal, see: https://stackoverflow.com/questions/1813872/running-a-process-in-pythonw-with-popen-without-a-console)
                startupinfo = STARTUPINFO()
                startupinfo.dwFlags |= STARTF_USESHOWWINDOW
                with Popen(ffmpeg_cmd, stdout=PIPE, stderr=STDOUT, bufsize=1, universal_newlines=True, encoding='utf-8', startupinfo=startupinfo) as ffmpeg_proc:
                    for line in ffmpeg_proc.stdout:
                        self.logn('ffmpeg: ' + line)
            elif platform.system() in ("Darwin", "Linux"):
                with Popen(ffmpeg_cmd, stdout=PIPE, stderr=STDOUT, bufsize=1, universal_newlines=True, encoding='utf-8') as ffmpeg_proc:
                    for line in ffmpeg_proc.stdout:
                        self.logn('ffmpeg: ' + line)
            
            if ffmpeg_proc.returncode > 0:
                raise Exception(f'FFmpeg conversion failed with return code {ffmpeg_proc.returncode}')
                
            # Verify output file was created
            if not os.path.exists(self.tmp_audio_file):
                raise Exception('Output file was not created by FFmpeg')
                
        except Exception as e:
            self.logn(f'Error converting audio file {self.audio_file}: {str(e)}', 'error')
            raise Exception(t('err_converting_audio'))

    def transcription_worker(self):
        # This is the main function where all the magic happens
        # We put this in a seperate thread so that it does not block the main ui

        proc_start_time = datetime.datetime.now()
        self.cancel = False

        # Show the stop button
        self.start_button.pack_forget() # hide
        self.stop_button.pack(padx=[20, 0], pady=[20,30], expand=False, fill='x', anchor='sw')
        
        # Show the progress bar
        # self.progress_bar.set(0)
        # self.progress_bar.pack(padx=[10,10], pady=[10,10], expand=True, fill='x', anchor='sw', side='left')
        # self.progress_bar.pack(padx=[0,10], pady=[10,25], expand=True, fill='x', anchor='sw', side='left')
        # self.progress_bar.pack(padx=20, pady=[10,20], expand=True, fill='both')

        tmpdir = TemporaryDirectory('noScribe')
        self.tmp_audio_file = os.path.join(tmpdir.name, 'tmp_audio.wav')

        try:
            # collect all the options
            option_info = ''

            if self.input_mode == 'single':
                if self.audio_file == '':
                    self.logn(t('err_no_audio_file'), 'error')
                    tk.messagebox.showerror(title='noScribe', message=t('err_no_audio_file'))
                    return
            else:  # multi-file mode
                if len(self.audio_files) == 0:
                    self.logn('Error: Please add at least one audio file.', 'error')
                    tk.messagebox.showerror(title='noScribe', message='Error: Please add at least one audio file.')
                    return

            if self.transcript_file == '':
                self.logn(t('err_no_transcript_file'), 'error')
                tk.messagebox.showerror(title='noScribe', message=t('err_no_transcript_file'))
                return

            self.my_transcript_file = self.transcript_file
            self.file_ext = os.path.splitext(self.my_transcript_file)[1][1:]

            # create log file
            if not os.path.exists(f'{config_dir}/log'):
                os.makedirs(f'{config_dir}/log')
            self.log_file = open(f'{config_dir}/log/{Path(self.my_transcript_file).stem}.log', 'w', encoding="utf-8")

            # options for faster-whisper
            self.whisper_beam_size = get_config('whisper_beam_size', 1)
            self.logn(f'whisper beam size: {self.whisper_beam_size}', where='file')

            self.whisper_temperature = get_config('whisper_temperature', 0.0)
            self.logn(f'whisper temperature: {self.whisper_temperature}', where='file')

            self.whisper_compute_type = get_config('whisper_compute_type', 'default')
            self.logn(f'whisper compute type: {self.whisper_compute_type}', where='file')

            self.timestamp_interval = get_config('timestamp_interval', 60_000) # default: add a timestamp every minute
            self.logn(f'timestamp_interval: {self.timestamp_interval}', where='file')

            self.timestamp_color = get_config('timestamp_color', '#78909C') # default: light gray/blue
            self.logn(f'timestamp_color: {self.timestamp_color}', where='file')

            # get UI settings
            val = self.entry_start.get()
            if val == '':
                self.start = 0
            else:
                self.start = millisec(val)
                option_info += f'{t("label_start")} {val} | ' 

            val = self.entry_stop.get()
            if val == '':
                self.stop = '0'
            else:
                self.stop = millisec(val)
                option_info += f'{t("label_stop")} {val} | '          
            
            sel_whisper_model = self.option_menu_whisper_model.get()
            if sel_whisper_model in self.whisper_model_paths.keys():
                self.whisper_model = self.whisper_model_paths[sel_whisper_model]
            else:                
                raise FileNotFoundError(f"The whisper model '{sel_whisper_model}' does not exist.")

            option_info += f'{t("label_whisper_model")} {sel_whisper_model} | '

            self.language_name = self.option_menu_language.get()
            option_info += f'{t("label_language")} {self.language_name} ({languages[self.language_name]}) | '

            self.speaker_detection = self.option_menu_speaker.get()
            option_info += f'{t("label_speaker")} {self.speaker_detection} | '

            self.overlapping = self.check_box_overlapping.get()
            option_info += f'{t("label_overlapping")} {self.overlapping} | '

            self.timestamps = self.check_box_timestamps.get()
            option_info += f'{t("label_timestamps")} {self.timestamps} | '
            
            self.disfluencies = self.check_box_disfluencies.get()
            option_info += f'{t("label_disfluencies")} {self.disfluencies} | '

            self.pause = self.option_menu_pause._values.index(self.option_menu_pause.get())
            option_info += f'{t("label_pause")} {self.pause}'

            self.pause_marker = get_config('pause_seconds_marker', '.') # Default to . if marker not in config

            # Default to True if auto save not in config or invalid value
            self.auto_save = False if get_config('auto_save', 'True') == 'False' else True 
            
            # Open the finished transript in the editor automatically?
            self.auto_edit_transcript = get_config('auto_edit_transcript', 'True')
            
            # Check for invalid vtt options
            if self.file_ext == 'vtt' and (self.pause > 0 or self.overlapping or self.timestamps):
                self.logn()
                self.logn(t('err_vtt_invalid_options'), 'error')
                self.pause = 0
                self.overlapping = False
                self.timestamps = False           

            if platform.system() == "Darwin": # = MAC
                # if (platform.mac_ver()[0] >= '12.3' and
                #     # torch.backends.mps.is_built() and # not necessary since depends on packaged PyTorch
                #     torch.backends.mps.is_available()):
                # Default to mps on 12.3 and newer, else cpu
                xpu = get_config('pyannote_xpu', 'mps' if platform.mac_ver()[0] >= '12.3' else 'cpu')
                self.pyannote_xpu = 'mps' if xpu == 'mps' else 'cpu'
            elif platform.system() in ('Windows', 'Linux'):
                # Use cuda if available and not set otherwise in config.yml, fallback to cpu: 
                cuda_available = torch.cuda.is_available() and get_cuda_device_count() > 0
                xpu = get_config('pyannote_xpu', 'cuda' if cuda_available else 'cpu')
                self.pyannote_xpu = 'cuda' if xpu == 'cuda' else 'cpu'
                whisper_xpu = get_config('whisper_xpu', 'cuda' if cuda_available else 'cpu')
                self.whisper_xpu = 'cuda' if whisper_xpu == 'cuda' else 'cpu'
            else:
                raise Exception('Platform not supported yet.')

            # log CPU capabilities
            self.logn("=== CPU FEATURES ===", where="file")
            if platform.system() == 'Windows':
                self.logn("System: Windows", where="file")
                for key, value in cpufeature.CPUFeature.items():
                    self.logn('    {:24}: {}'.format(key, value), where="file")
            elif platform.system() == "Darwin": # = MAC
                self.logn(f"System: MAC {platform.machine()}", where="file")
                if platform.mac_ver()[0] >= '12.3': # MPS needs macOS 12.3+
                    if config['pyannote_xpu'] == 'mps':
                        self.logn("macOS version >= 12.3:\nUsing MPS (with PYTORCH_ENABLE_MPS_FALLBACK enabled)", where="file")
                    elif config['pyannote_xpu'] == 'cpu':
                        self.logn("macOS version >= 12.3:\nUser selected to use CPU (results will be better, but you might wanna make yourself a coffee)", where="file")
                    else:
                        self.logn("macOS version >= 12.3:\nInvalid option for 'pyannote_xpu' in config.yml (should be 'mps' or 'cpu')\nYou might wanna change this\nUsing MPS anyway (with PYTORCH_ENABLE_MPS_FALLBACK enabled)", where="file")
                else:
                    self.logn("macOS version < 12.3:\nMPS not available: Using CPU\nPerformance might be poor\nConsider updating macOS, if possible", where="file")

            try:

                #-------------------------------------------------------
                # 1) Convert Audio or Handle Multi-File

                try:
                    if self.input_mode == 'multi':
                        # Multi-file mode: transcribe each file separately
                        self.logn()
                        self.logn('Merging audio files...', 'highlight')
                        
                        # Transcribe multiple files
                        all_segments = self.transcribe_multi_files(tmpdir)
                        
                        self.logn('Audio merge finished')
                        
                        # Create transcript from merged segments
                        self.create_transcript_from_segments(all_segments)
                        
                        self.logn()
                        self.logn(t('transcription_finished'), 'highlight')
                        self.log(t('transcription_saved'))
                        self.logn(self.transcript_file, link=f'file://{self.transcript_file}')
                        
                        # Log duration
                        proc_time = datetime.datetime.now() - proc_start_time
                        proc_seconds = "{:02d}".format(int(proc_time.total_seconds() % 60))
                        proc_time_str = f'{int(proc_time.total_seconds() // 60)}:{proc_seconds}' 
                        self.logn(f'Duration: {proc_time_str} minutes')
                        
                        # Auto open transcript in editor
                        if (self.auto_edit_transcript == 'True') and (self.file_ext == 'html'):
                            self.launch_editor(self.transcript_file)
                        
                        return  # Exit early for multi-file mode
                    
                    # Single file mode (existing logic)
                    self.logn()
                    self.logn(t('start_audio_conversion'), 'highlight')
                
                    if int(self.stop) > 0: # transcribe only part of the audio
                        end_pos_cmd = f'-to {self.stop}ms'
                    else: # tranbscribe until the end
                        end_pos_cmd = ''

                    # Use the same fixed audio conversion logic as multi-file mode
                    self.convert_audio_file_single_mode()
                    self.logn(t('audio_conversion_finished'))
                    self.set_progress(1, 50)
                except Exception as e:
                    self.logn(t('err_converting_audio'), 'error')
                    traceback_str = traceback.format_exc()
                    self.logn(e, 'error', tb=traceback_str)
                    return

                #-------------------------------------------------------
                # 2) Speaker identification (diarization) with pyannote

                # Helper Functions:

                def overlap_len(ss_start, ss_end, ts_start, ts_end):
                    # ss...: speaker segment start and end in milliseconds (from pyannote)
                    # ts...: transcript segment start and end (from whisper.cpp)
                    # returns overlap percentage, i.e., "0.8" = 80% of the transcript segment overlaps with the speaker segment from pyannote  
                    if ts_end < ss_start: # no overlap, ts is before ss
                        return None

                    if ts_start > ss_end: # no overlap, ts is after ss
                        return 0.0

                    ts_len = ts_end - ts_start
                    if ts_len <= 0:
                        return None

                    # ss & ts have overlap
                    overlap_start = max(ss_start, ts_start) # Whichever starts later
                    overlap_end = min(ss_end, ts_end) # Whichever ends sooner

                    ol_len = overlap_end - overlap_start + 1
                    return ol_len / ts_len

                def find_speaker(diarization, transcript_start, transcript_end) -> str:
                    # Looks for the shortest segment in diarization that has at least 80% overlap 
                    # with transcript_start - trancript_end.  
                    # Returns the speaker name if found.
                    # If only an overlap < 80% is found, this speaker name ist returned.
                    # If no overlap is found, an empty string is returned.
                    spkr = ''
                    overlap_found = 0
                    overlap_threshold = 0.8
                    segment_len = 0
                    is_overlapping = False

                    for segment in diarization:
                        t = overlap_len(segment["start"], segment["end"], transcript_start, transcript_end)
                        if t is None: # we are already after transcript_end
                            break

                        current_segment_len = segment["end"] - segment["start"] # Length of the current segment
                        current_segment_spkr = f'S{segment["label"][8:]}' # shorten the label: "SPEAKER_01" > "S01"

                        if overlap_found >= overlap_threshold: # we already found a fitting segment, compare length now
                            if (t >= overlap_threshold) and (current_segment_len < segment_len): # found a shorter (= better fitting) segment that also overlaps well
                                is_overlapping = True
                                overlap_found = t
                                segment_len = current_segment_len
                                spkr = current_segment_spkr
                        elif t > overlap_found: # no segment with good overlap yet, take this if the overlap is better then previously found 
                            overlap_found = t
                            segment_len = current_segment_len
                            spkr = current_segment_spkr
                        
                    if self.overlapping and is_overlapping:
                        return f"//{spkr}"
                    else:
                        return spkr

                # Start Diarization:

                if self.speaker_detection != 'none':
                    try:
                        self.logn()
                        self.logn(t('start_identifiying_speakers'), 'highlight')
                        self.logn(t('loading_pyannote'))
                        self.set_progress(1, 100)

                        diarize_output = os.path.join(tmpdir.name, 'diarize_out.yaml')
                        diarize_abspath = 'python ' + os.path.join(app_dir, 'diarize.py')
                        diarize_abspath_win = os.path.join(app_dir, '..', 'diarize.exe')
                        diarize_abspath_mac = os.path.join(app_dir, '..', 'MacOS', 'diarize')
                        diarize_abspath_lin = os.path.join(app_dir, '..', 'diarize')
                        if platform.system() == 'Windows' and os.path.exists(diarize_abspath_win):
                            diarize_abspath = diarize_abspath_win
                        elif platform.system() == 'Darwin' and os.path.exists(diarize_abspath_mac): # = MAC
                            diarize_abspath = diarize_abspath_mac
                        elif platform.system() == 'Linux' and os.path.exists(diarize_abspath_lin):
                            diarize_abspath = diarize_abspath_lin
                        diarize_cmd = f'{diarize_abspath} {self.pyannote_xpu} "{self.tmp_audio_file}" "{diarize_output}" {self.speaker_detection}'
                        diarize_env = None
                        if self.pyannote_xpu == 'mps':
                            diarize_env = os.environ.copy()
                            diarize_env["PYTORCH_ENABLE_MPS_FALLBACK"] = str(1) # Necessary since some operators are not implemented for MPS yet.
                        self.logn(diarize_cmd, where='file')

                        if platform.system() == 'Windows':
                            # (supresses the terminal, see: https://stackoverflow.com/questions/1813872/running-a-process-in-pythonw-with-popen-without-a-console)
                            startupinfo = STARTUPINFO()
                            startupinfo.dwFlags |= STARTF_USESHOWWINDOW
                        elif platform.system() in ('Darwin', "Linux"): # = MAC
                            diarize_cmd = shlex.split(diarize_cmd)
                            startupinfo = None
                        else:
                            raise Exception('Platform not supported yet.')

                        with Popen(diarize_cmd,
                                   stdout=PIPE,
                                   stderr=STDOUT,
                                   encoding='UTF-8',
                                   startupinfo=startupinfo,
                                   env=diarize_env,
                                   close_fds=True) as pyannote_proc:
                            for line in pyannote_proc.stdout:
                                if self.cancel:
                                    pyannote_proc.kill()
                                    raise Exception(t('err_user_cancelation')) 
                                print(line)
                                if line.startswith('progress '):
                                    progress = line.split()
                                    step_name = progress[1]
                                    progress_percent = int(progress[2])
                                    self.logr(f'{step_name}: {progress_percent}%')                       
                                    if step_name == 'segmentation':
                                        self.set_progress(2, progress_percent * 0.3)
                                    elif step_name == 'embeddings':
                                        self.set_progress(2, 30 + (progress_percent * 0.7))
                                elif line.startswith('error '):
                                    self.logn('PyAnnote error: ' + line[5:], 'error')
                                elif line.startswith('log: '):
                                    self.logn('PyAnnote ' + line, where='file')
                                    if line.strip() == "log: 'pyannote_xpu: cpu' was set.": # The string needs to be the same as in diarize.py `print("log: 'pyannote_xpu: cpu' was set.")`.
                                        self.pyannote_xpu = 'cpu'
                                        config['pyannote_xpu'] = 'cpu'

                        if pyannote_proc.returncode > 0:
                            raise Exception('')

                        # load diarization results
                        with open(diarize_output, 'r') as file:
                            diarization = yaml.safe_load(file)

                        # write segments to log file 
                        for segment in diarization:
                            line = f'{ms_to_str(self.start + segment["start"], include_ms=True)} - {ms_to_str(self.start + segment["end"], include_ms=True)} {segment["label"]}'
                            self.logn(line, where='file')

                        self.logn()

                    except Exception as e:
                        self.logn(t('err_identifying_speakers'), 'error')
                        traceback_str = traceback.format_exc()
                        self.logn(e, 'error', tb=traceback_str)
                        return

                #-------------------------------------------------------
                # 3) Transcribe with faster-whisper

                self.logn()
                self.logn(t('start_transcription'), 'highlight')
                self.logn(t('loading_whisper'))

                # prepare transcript html
                d = AdvancedHTMLParser.AdvancedHTMLParser()
                d.parseStr(default_html)                

                # add audio file path:
                tag = d.createElement("meta")
                tag.name = "audio_source"
                tag.content = self.audio_file
                d.head.appendChild(tag)

                # add app version:
                """ # removed because not really necessary
                tag = d.createElement("meta")
                tag.name = "noScribe_version"
                tag.content = app_version
                d.head.appendChild(tag)
                """

                #add WordSection1 (for line numbers in MS Word) as main_body
                main_body = d.createElement('div')
                main_body.addClass('WordSection1')
                d.body.appendChild(main_body)

                # header               
                p = d.createElement('p')
                p.setStyle('font-weight', '600')
                p.appendText(Path(self.audio_file).stem) # use the name of the audio file (without extension) as the title
                main_body.appendChild(p)

                # subheader
                p = d.createElement('p')
                s = d.createElement('span')
                s.setStyle('color', '#909090')
                s.setStyle('font-size', '0.8em')
                s.appendText(t('doc_header', version=app_version))
                br = d.createElement('br')
                s.appendChild(br)

                s.appendText(t('doc_header_audio', file=self.audio_file))
                br = d.createElement('br')
                s.appendChild(br)

                s.appendText(f'({html.escape(option_info)})')

                p.appendChild(s)
                main_body.appendChild(p)

                p = d.createElement('p')
                main_body.appendChild(p)

                speaker = ''
                prev_speaker = ''
                self.last_auto_save = datetime.datetime.now()

                def save_doc():
                    txt = ''
                    if self.file_ext == 'html':
                        txt = d.asHTML()
                    elif self.file_ext == 'txt':
                        txt = html_to_text(d)
                    elif self.file_ext == 'vtt':
                        txt = html_to_webvtt(d, self.audio_file)
                    else:
                        raise TypeError(f'Invalid file type "{self.file_ext}".')
                    try:
                        if txt != '':
                            with open(self.my_transcript_file, 'w', encoding="utf-8") as f:
                                f.write(txt)
                                f.flush()
                            self.last_auto_save = datetime.datetime.now()
                    except Exception as e:
                        # Error while saving - could be read-only filesystem, file already open, or other issues
                        transcript_path = Path(self.my_transcript_file)
                        
                        # Check if it's a read-only filesystem error (errno 30)
                        if hasattr(e, 'errno') and e.errno == 30:
                            # Read-only filesystem - redirect to user's Desktop
                            desktop_path = Path.home() / 'Desktop'
                            fallback_filename = f'{transcript_path.stem}_transcript.{self.file_ext}'
                            self.my_transcript_file = str(desktop_path / fallback_filename)
                            
                            # Ensure unique filename on Desktop
                            counter = 1
                            while os.path.exists(self.my_transcript_file):
                                fallback_filename = f'{transcript_path.stem}_transcript_{counter}.{self.file_ext}'
                                self.my_transcript_file = str(desktop_path / fallback_filename)
                                counter += 1
                                
                            self.logn()
                            self.logn(f'Cannot save to read-only location. Saving to Desktop: {self.my_transcript_file}', 'error')
                        else:
                            # Other error (maybe file already open in Word) - try alternative filename in same directory
                            self.my_transcript_file = f'{transcript_path.parent}/{transcript_path.stem}_1.{self.file_ext}'
                            if os.path.exists(self.my_transcript_file):
                                # Alternative filename also exists - try Desktop as last resort
                                desktop_path = Path.home() / 'Desktop'
                                fallback_filename = f'{transcript_path.stem}_transcript_rescue.{self.file_ext}'
                                self.my_transcript_file = str(desktop_path / fallback_filename)
                                
                                counter = 1
                                while os.path.exists(self.my_transcript_file):
                                    fallback_filename = f'{transcript_path.stem}_transcript_rescue_{counter}.{self.file_ext}'
                                    self.my_transcript_file = str(desktop_path / fallback_filename)
                                    counter += 1
                        
                        try:
                            with open(self.my_transcript_file, 'w', encoding="utf-8") as f:
                                f.write(txt)
                                f.flush()
                            self.logn()
                            self.logn(t('rescue_saving', file=self.my_transcript_file), 'error', link=f'file://{self.my_transcript_file}')
                            self.last_auto_save = datetime.datetime.now()
                        except Exception as e2:
                            # Last resort failed too
                            raise Exception(f'Failed to save transcript to any location. Original error: {str(e)}, Rescue error: {str(e2)}')

                try:
                    from faster_whisper import WhisperModel
                    if platform.system() == "Darwin": # = MAC
                        whisper_device = 'auto'
                    elif platform.system() in ('Windows', 'Linux'):
                        whisper_device = 'cpu'
                        whisper_device = self.whisper_xpu
                    else:
                        raise Exception('Platform not supported yet.')
                    model = WhisperModel(self.whisper_model,
                                         device=whisper_device,  
                                         cpu_threads=number_threads, 
                                         compute_type=self.whisper_compute_type, 
                                         local_files_only=True)
                    self.logn('model loaded', where='file')

                    if self.cancel:
                        raise Exception(t('err_user_cancelation')) 

                    multilingual = False
                    if self.language_name == 'Multilingual':
                        multilingual = True
                        whisper_lang = None
                    elif self.language_name == 'Auto':
                        whisper_lang = None
                    else:
                        whisper_lang = languages[self.language_name]
                    
                    # VAD 
                     
                    try:
                        self.vad_threshold = float(config['voice_activity_detection_threshold'])
                    except:
                        config['voice_activity_detection_threshold'] = '0.5'
                        self.vad_threshold = 0.5                     

                    sampling_rate = model.feature_extractor.sampling_rate
                    audio = decode_audio(self.tmp_audio_file, sampling_rate=sampling_rate)
                    duration = audio.shape[0] / sampling_rate
                    
                    self.logn('Voice Activity Detection')
                    try:
                        vad_parameters = VadOptions(min_silence_duration_ms=1000, 
                                                threshold=self.vad_threshold,
                                                speech_pad_ms=0)
                    except TypeError:
                        # parameter threshold was temporarily renamed to 'onset' in pyannote 3.1:  
                        vad_parameters = VadOptions(min_silence_duration_ms=1000, 
                                                onset=self.vad_threshold,
                                                speech_pad_ms=0)
                    speech_chunks = get_speech_timestamps(audio, vad_parameters)
                    
                    def adjust_for_pause(segment):
                        """Adjusts start and end of segment if it falls into a pause 
                        identified by the VAD"""
                        pause_extend = 0.2  # extend the pauses by 200ms to make the detection more robust
                        
                        # iterate through the pauses and adjust segment boundaries accordingly
                        for i in range(0, len(speech_chunks)):
                            pause_start = (speech_chunks[i]['end'] / sampling_rate) - pause_extend
                            if i == (len(speech_chunks) - 1): 
                                pause_end = duration + pause_extend # last segment, pause till the end
                            else:
                                pause_end = (speech_chunks[i+1]['start']  / sampling_rate) + pause_extend
                            
                            if pause_start > segment.end:
                                break  # we moved beyond the segment, stop going further
                            if segment.start > pause_start and segment.start < pause_end:
                                segment.start = pause_end - pause_extend
                            if segment.end > pause_start and segment.end < pause_end:
                                segment.end = pause_start + pause_extend
                        
                        return segment
                    
                    # transcribe
                    
                    if self.cancel:
                        raise Exception(t('err_user_cancelation')) 

                    vad_parameters.speech_pad_ms = 400

                    # detect language                    
                    if self.language_name == 'auto':
                        language, language_probability, all_language_probs = model.detect_language(
                            audio,
                            vad_filter=True,
                            vad_parameters=vad_parameters
                        )
                        self.language = language
                        self.logn("Detected language '%s' with probability %f" % (language, language_probability))

                    if self.disfluencies:                    
                        try:
                            with open(os.path.join(app_dir, 'prompt.yml'), 'r', encoding='utf-8') as file:
                                prompts = yaml.safe_load(file)
                        except:
                            prompts = {}
                        self.prompt = prompts.get(languages[self.language_name], '') # Fetch language prompt, default to empty string
                    else:
                        self.prompt = ''
                    
                    del audio
                    gc.collect()
                    
                    segments, info = model.transcribe(
                        self.tmp_audio_file, # audio, 
                        language=whisper_lang,
                        multilingual=multilingual, 
                        beam_size=5, 
                        #temperature=self.whisper_temperature, 
                        word_timestamps=True, 
                        #initial_prompt=self.prompt,
                        hotwords=self.prompt, 
                        vad_filter=True,
                        vad_parameters=vad_parameters,
                        # length_penalty=0.5
                    )

                    if self.cancel:
                        raise Exception(t('err_user_cancelation')) 

                    self.logn(t('start_transcription'))
                    self.logn()

                    last_segment_end = 0
                    last_timestamp_ms = 0
                    first_segment = True

                    for segment in segments:
                        # check for user cancelation
                        if self.cancel:
                            if self.auto_save:
                                save_doc()
                                self.logn()
                                self.log(t('transcription_saved'))
                                self.logn(self.my_transcript_file, link=f'file://{self.my_transcript_file}')

                            raise Exception(t('err_user_cancelation')) 

                        segment = adjust_for_pause(segment)

                        # get time of the segment in milliseconds
                        start = round(segment.start * 1000.0)
                        end = round(segment.end * 1000.0)
                        # if we skipped a part at the beginning of the audio we have to add this here again, otherwise the timestaps will not match the original audio:
                        orig_audio_start = self.start + start
                        orig_audio_end = self.start + end

                        if self.timestamps:
                            ts = ms_to_str(orig_audio_start)
                            ts = f'[{ts}]'

                        # check for pauses and mark them in the transcript
                        if (self.pause > 0) and (start - last_segment_end >= self.pause * 1000): # (more than x seconds with no speech)
                            pause_len = round((start - last_segment_end)/1000)
                            if pause_len >= 60: # longer than 60 seconds
                                pause_str = ' ' + t('pause_minutes', minutes=round(pause_len/60))
                            elif pause_len >= 10: # longer than 10 seconds
                                pause_str = ' ' + t('pause_seconds', seconds=pause_len)
                            else: # less than 10 seconds
                                pause_str = ' (' + (self.pause_marker * pause_len) + ')'

                            if first_segment:
                                pause_str = pause_str.lstrip() + ' '

                            orig_audio_start_pause = self.start + last_segment_end
                            orig_audio_end_pause = self.start + start
                            a = d.createElement('a')
                            a.name = f'ts_{orig_audio_start_pause}_{orig_audio_end_pause}_{speaker}'
                            a.appendText(pause_str)
                            p.appendChild(a)
                            self.log(pause_str)
                            if first_segment:
                                self.logn()
                                self.logn()
                        last_segment_end = end

                        # write text to the doc
                        # diarization (speaker detection)?
                        seg_text = segment.text
                        seg_html = html.escape(seg_text)

                        if self.speaker_detection != 'none':
                            new_speaker = find_speaker(diarization, start, end)
                            if (speaker != new_speaker) and (new_speaker != ''): # speaker change
                                if new_speaker[:2] == '//': # is overlapping speech, create no new paragraph
                                    prev_speaker = speaker
                                    speaker = new_speaker
                                    seg_text = f' {speaker}:{seg_text}'
                                    seg_html = html.escape(seg_text)                                
                                elif (speaker[:2] == '//') and (new_speaker == prev_speaker): # was overlapping speech and we are returning to the previous speaker 
                                    speaker = new_speaker
                                    seg_text = f'//{seg_text}'
                                    seg_html = html.escape(seg_text)
                                else: # new speaker, not overlapping
                                    if speaker[:2] == '//': # was overlapping speech, mark the end
                                        last_elem = p.lastElementChild
                                        if last_elem:
                                            last_elem.appendText('//')
                                        else:
                                            p.appendText('//')
                                        self.log('//')
                                    p = d.createElement('p')
                                    main_body.appendChild(p)
                                    if not first_segment:
                                        self.logn()
                                        self.logn()
                                    speaker = new_speaker
                                    # add timestamp
                                    if self.timestamps:
                                        seg_html = f'{speaker}: <span style="color: {self.timestamp_color}" >{ts}</span>{html.escape(seg_text)}'
                                        seg_text = f'{speaker}: {ts}{seg_text}'
                                        last_timestamp_ms = start
                                    else:
                                        if self.file_ext != 'vtt': # in vtt files, speaker names are added as special voice tags so skip this here
                                            seg_text = f'{speaker}:{seg_text}'
                                            seg_html = html.escape(seg_text)
                                        else:
                                            seg_html = html.escape(seg_text).lstrip()
                                            seg_text = f'{speaker}:{seg_text}'
                                        
                            else: # same speaker
                                if self.timestamps:
                                    if (start - last_timestamp_ms) > self.timestamp_interval:
                                        seg_html = f' <span style="color: {self.timestamp_color}" >{ts}</span>{html.escape(seg_text)}'
                                        seg_text = f' {ts}{seg_text}'
                                        last_timestamp_ms = start
                                    else:
                                        seg_html = html.escape(seg_text)

                        else: # no speaker detection
                            if self.timestamps and (first_segment or (start - last_timestamp_ms) > self.timestamp_interval):
                                seg_html = f' <span style="color: {self.timestamp_color}" >{ts}</span>{html.escape(seg_text)}'
                                seg_text = f' {ts}{seg_text}'
                                last_timestamp_ms = start
                            else:
                                seg_html = html.escape(seg_text)
                            # avoid leading whitespace in first paragraph
                            if first_segment:
                                seg_text = seg_text.lstrip()
                                seg_html = seg_html.lstrip()

                        # Mark confidence level (not implemented yet in html)
                        # cl_level = round((segment.avg_logprob + 1) * 10)
                        # TODO: better cl_level for words based on https://github.com/Softcatala/whisper-ctranslate2/blob/main/src/whisper_ctranslate2/transcribe.py
                        # if cl_level > 0:
                        #     r.style = d.styles[f'noScribe_cl{cl_level}']

                        # Create bookmark with audio timestamps start to end and add the current segment.
                        # This way, we can jump to the according audio position and play it later in the editor.
                        a_html = f'<a name="ts_{orig_audio_start}_{orig_audio_end}_{speaker}" >{seg_html}</a>'
                        a = d.createElementFromHTML(a_html)
                        p.appendChild(a)

                        self.log(seg_text)

                        first_segment = False

                        # auto save
                        if self.auto_save:
                            if (datetime.datetime.now() - self.last_auto_save).total_seconds() > 20:
                                save_doc()

                        progr = round((segment.end/info.duration) * 100)
                        self.set_progress(3, progr)

                    save_doc()
                    self.logn()
                    self.logn()
                    self.logn(t('transcription_finished'), 'highlight')
                    if self.transcript_file != self.my_transcript_file: # used alternative filename because saving under the initial name failed
                        self.log(t('rescue_saving'))
                        self.logn(self.my_transcript_file, link=f'file://{self.my_transcript_file}')
                    else:
                        self.log(t('transcription_saved'))
                        self.logn(self.my_transcript_file, link=f'file://{self.my_transcript_file}')
                    # log duration of the whole process
                    proc_time = datetime.datetime.now() - proc_start_time
                    proc_seconds = "{:02d}".format(int(proc_time.total_seconds() % 60))
                    proc_time_str = f'{int(proc_time.total_seconds() // 60)}:{proc_seconds}' 
                    self.logn(t('trancription_time', duration=proc_time_str)) 

                    # auto open transcript in editor
                    if (self.auto_edit_transcript == 'True') and (self.file_ext == 'html'):
                        self.launch_editor(self.my_transcript_file)
                
                except Exception as e:
                    self.logn()
                    self.logn(t('err_transcription'), 'error')
                    traceback_str = traceback.format_exc()
                    self.logn(e, 'error', tb=traceback_str)
                    return

            finally:
                self.log_file.close()
                self.log_file = None

        except Exception as e:
            self.logn(t('err_options'), 'error')
            traceback_str = traceback.format_exc()
            self.logn(e, 'error', tb=traceback_str)
            return

        finally:
            # hide the stop button
            self.stop_button.pack_forget() # hide
            self.start_button.pack(padx=[20, 0], pady=[20,30], expand=False, fill='x', anchor='sw')

            # hide progress
            self.set_progress(0, 0)
            
    def button_start_event(self):
        wkr = Thread(target=self.transcription_worker)
        wkr.start()
        #while wkr.is_alive():
        #    self.update()
        #    time.sleep(0.1)
    
    # End main function Button Start        
    ################################################################################################

    def button_stop_event(self):
        if tk.messagebox.askyesno(title='noScribe', message=t('transcription_canceled')) == True:
            self.logn()
            self.logn(t('start_canceling'))
            self.update()
            self.cancel = True

    def on_closing(self):
        # (see: https://stackoverflow.com/questions/111155/how-do-i-handle-the-window-close-event-in-tkinter)
        #if messagebox.askokcancel("Quit", "Do you want to quit?"):
        try:
            # remember some settings for the next run
            config['last_language'] = self.option_menu_language.get()
            config['last_speaker'] = self.option_menu_speaker.get()
            config['last_whisper_model'] = self.option_menu_whisper_model.get()
            config['last_pause'] = self.option_menu_pause.get()
            config['last_overlapping'] = self.check_box_overlapping.get()
            config['last_timestamps'] = self.check_box_timestamps.get()
            config['last_disfluencies'] = self.check_box_disfluencies.get()

            save_config()
        finally:
            self.destroy()

    def merge_word_segments_to_utterances(self, word_segments, max_utterance_duration=60, min_intermittent_duration=0.8):
        """
        Merge word-level segments into speaker utterances using transview's approach.
        This handles timeline accuracy while creating readable transcript segments.
        """
        if not word_segments:
            return []
        
        merged_utterances = []
        current_utterance = None
        
        for segment in word_segments:
            if current_utterance is None:
                # Start first utterance
                current_utterance = {
                    'start': segment['start'],
                    'end': segment['end'],
                    'text': segment['text'],
                    'speaker': segment['speaker']
                }
            elif (segment['speaker'] == current_utterance['speaker'] and 
                  (current_utterance['end'] - current_utterance['start']) < max_utterance_duration):
                # Same speaker and utterance not too long - merge
                current_utterance['end'] = segment['end']
                current_utterance['text'] += ' ' + segment['text']
            else:
                # Speaker change or utterance too long - finalize current and start new
                merged_utterances.append(current_utterance)
                current_utterance = {
                    'start': segment['start'],
                    'end': segment['end'],
                    'text': segment['text'],
                    'speaker': segment['speaker']
                }
        
        # Add final utterance
        if current_utterance:
            merged_utterances.append(current_utterance)
        
        # Handle intermittent speaker changes (short interruptions)
        final_utterances = self.handle_intermittent_speaker_changes(merged_utterances, min_intermittent_duration)
        
        return final_utterances

    def handle_intermittent_speaker_changes(self, utterances, min_duration=0.8):
        """
        Handle short intermittent speaker changes that should be merged back.
        Based on transview's _helper_flatten_intermittent_segments approach.
        """
        if len(utterances) < 4:
            return utterances
        
        result = utterances[:3]  # Keep first 3 utterances as-is
        
        for i, utterance in enumerate(utterances[3:], 3):
            # Check if we have a pattern of intermittent short segments
            prev3 = result[-3]  # 3 positions back
            prev2 = result[-2]  # 2 positions back  
            prev1 = result[-1]  # 1 position back
            current = utterance
            
            # Duration check for intermittent segments
            prev2_duration = prev2['end'] - prev2['start']
            prev1_duration = prev1['end'] - prev1['start']
            
            if (prev2_duration < min_duration and 
                prev1_duration < min_duration and
                prev3['speaker'] == prev1['speaker'] and
                prev2['speaker'] == current['speaker'] and
                prev1['speaker'] != current['speaker']):
                
                # Merge intermittent segments back to their primary speakers
                # Merge prev1 back to prev3 (same speaker)
                prev3['end'] = prev1['end']
                prev3['text'] += ' ' + prev1['text']
                
                # Merge current back to prev2 (same speaker)
                prev2['end'] = current['end']
                prev2['text'] += ' ' + current['text']
                
                # Remove the merged segment (prev1)
                result.pop(-1)
            else:
                result.append(current)
        
        return result

if __name__ == "__main__":

    app = App()

    app.mainloop()
