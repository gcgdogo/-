#coding:UTF-8
from __future__ import division,print_function,absolute_import

###########################################
#利用装饰器直接就把设置命令执行了
import kivy
import sys
print("[设置 Kivy]<<开始>>#######################")
print("[设置 Kivy]" + '< kivy.__version__ > :: ' + repr(kivy.__version__))
def Config_Execute(*args):
    def Config_Function(func):
        if kivy.__version__ in args:
            print("[设置 Kivy]" + repr(func) + " :: " + "执行 [■]")
            func()
        else:
            print("[设置 Kivy]" + repr(func) + " :: " + "跳过 [ ]")
        return func
    return Config_Function

###########################################


#字体设置部分，暂时用的是一个字体，所以没有加粗和斜体，等以后有需要再研究怎么设置吧
@Config_Execute('1.9.0')
def Config001009000_ChineseFont():
    from kivy.core import core_select_lib
    from kivy.setupconfig import USE_SDL2
    #Load the appropriate provider
    label_libs = []
    if USE_SDL2:
        label_libs += [('sdl2', 'text_sdl2', 'LabelSDL2')]
    else:
        label_libs += [('pygame', 'text_pygame', 'LabelPygame')]
    label_libs += [('pil', 'text_pil', 'LabelPIL')]
    Label = core_select_lib('text', label_libs)
    Label.register('DroidSans', 'DroidSansFallback.ttf', 'DroidSansFallback.ttf', 'DroidSansFallback.ttf', 'DroidSansFallback.ttf')

@Config_Execute('1.11.1')
def Config001011001_ChineseFont():
    from kivy.config import Config
    #Config.set('graphics', 'default_font', ['DroidSansFallback', 'DroidSansFallback.ttf', 'DroidSansFallback.ttf']) #旧版的设置位置
    Config.set('kivy', 'default_font', ['DroidSansFallback', 'DroidSansFallback.ttf', 'DroidSansFallback.ttf']) #新版的设置位置
    ##################################################
    #Config.write()      轻易别write write之后可能就恢复不回来了
    ###################################################


print("[设置 Kivy]<<结束>>#######################")